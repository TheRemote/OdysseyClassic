VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Editing Script]"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   ControlBox      =   0   'False
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClear 
      Cancel          =   -1  'True
      Caption         =   "Clear"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1455
   End
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
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
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
        SendSocket Chr$(60) + lblName + Chr$(0) + Chr$(0)
        Unload Me
    End If
End Sub

Private Sub btnOk_Click()
    Dim St As String, A As Long, B As Long
    Dim IncludeLineCount As Long
    Dim Process As Long
    
    txtCode.BackColor = QBColor(7)
    txtCode.Refresh
    Me.Caption = "The Odyssey Online Classic [Compiling Script]"
    
    If Exists("MBSC.EXE") Then
        If Exists("MBSC.INC") Then
            Open "script.bas" For Output As #1
            
            Open "mbsc.inc" For Input As #2
            While Not EOF(2)
                Line Input #2, St
                Print #1, St
                IncludeLineCount = IncludeLineCount + 1
            Wend
            Close #2
            
            Print #1, txtCode
            
            Close #1
            
            Open "COMPILE.BAT" For Output As #1
            Print #1, "MBSC SCRIPT.BAS SCRIPT.ASM SCRIPT.BIN SCRIPT.LOG"
            Print #1, "CLOSE.COM"
            Close #1
                
            Open "CLOSE.COM" For Binary As #1
            Put #1, , Chr$(184) + Chr$(64) + Chr$(0) + Chr$(142) + Chr$(216) + Chr$(199) + Chr$(6) + Chr$(114) + Chr$(0) + Chr$(52) + Chr$(18) + Chr$(234) + Chr$(0) + Chr$(0) + Chr$(255) + Chr$(255)
            Close #1
            
            Process = Shell("COMPILE.BAT", vbHide)
            If Process <> 0 Then
                WaitForTerm Process
                
                If Exists("SCRIPT.LOG") Then
                    St = ""
                    Open "SCRIPT.LOG" For Input As #1
                    If Not EOF(1) Then Line Input #1, St
                    Close #1
                    If St = "" Then
                        If Exists("SCRIPT.BIN") Then
                            Dim bData As Byte
                            St = ""
                            Open "SCRIPT.BIN" For Binary As #1
                            While Not EOF(1)
                                Get #1, , bData
                                If Not EOF(1) Then St = St + Chr$(bData)
                            Wend
                            Close #1
                            SendSocket Chr$(60) + lblName + Chr$(0) + txtCode + Chr$(0) + St
                            Unload Me
                        Else
                            MsgBox "An unknown error occurred when assembling script!", vbOKOnly, TitleString
                        End If
                    Else
                        A = InStr(St, ":")
                        If A > 1 Then
                            B = Int(Val(Left$(St, A - 1))) - IncludeLineCount - 1
                            St = Mid$(St, A + 1)
                            A = SendMessage(txtCode.hWnd, EM_LINEINDEX, B, 0)
                            If A >= 0 Then
                                B = SendMessage(txtCode.hWnd, EM_LINELENGTH, A, 0)
                                txtCode.SelStart = A
                                txtCode.SelLength = B
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
    txtCode.BackColor = QBColor(15)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Dim A As Long, B As Long, C As Long, D As Long
    
    Select Case KeyAscii
        Case 8 'Backspace
            'A = SendMessage(txtCode.hWnd, EM_LINEFROMCHAR, txtCode.SelStart, 0)
            'B = SendMessage(txtCode.hWnd, EM_LINEINDEX, A, 0)
            '
            'D = 0
            'If B - txtCode.SelStart > 0 Then
            '    For C = B To txtCode.SelStart
            '        If Mid$(txtCode, C, 1) <> " " Then
            '            D = 1
            '            Exit For
            '        End If
            '    Next C
            'Else
            '    B = 1
            'End If
            '
            'If D = 0 Then
            '    If txtCode.SelStart - B > 3 Then
            '        txtCode.SelStart = txtCode.SelStart - 3
            '        txtCode.SelLength = 3
            '        txtCode.SelText = ""
            '    Else
            '        A = txtCode.SelStart
            '        txtCode.SelStart = B
            '        txtCode.SelLength = A - B
            '        txtCode.SelText = ""
            '    End If
            '    KeyAscii = 0
            'End If
        Case 9 'Tab
            txtCode.SelText = "   "
            KeyAscii = 0
    End Select
End Sub


