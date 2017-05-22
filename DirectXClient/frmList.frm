VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic"
   ClientHeight    =   5775
   ClientLeft      =   390
   ClientTop       =   45
   ClientWidth     =   3225
   ControlBox      =   0   'False
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox lstList 
      Height          =   4545
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Dim sTmp() As String
    Dim TrueIndex As Integer
    If lstList.ListIndex >= 0 Then
        sTmp = Split(lstList.List(lstList.ListIndex), ":")
        TrueIndex = CInt(sTmp(0))
        Select Case ListEditMode
        Case modeObjects
            SendSocket Chr$(19) + DoubleChar(CLng(TrueIndex))
        Case modeMonsters
            SendSocket Chr$(20) + DoubleChar$(CLng(TrueIndex))
        Case modeNPCs
            SendSocket Chr$(50) + DoubleChar$(CLng(TrueIndex))
        Case modeHalls
            SendSocket Chr$(48) + Chr$(TrueIndex)
        Case modeMagic
            SendSocket Chr$(82) + DoubleChar$(CLng(TrueIndex))
        Case modeBans
            SendSocket Chr$(57) + Chr$(lstList.ItemData(lstList.ListIndex))
        Case modePrefix
            SendSocket Chr$(86) + Chr$(TrueIndex)
        Case modeSuffix
            SendSocket Chr$(88) + Chr$(TrueIndex)
        End Select
        Me.Hide
    End If
End Sub

Public Sub DrawList()
    If ListEditMode = modeBans Then Exit Sub

    lstList.Clear

    Dim A As Long, B As Long

    Select Case ListEditMode
    Case modeObjects
        For A = 1 To MaxObjects
            If Len(txtSearch) > 0 Then
                B = InStr(UCase$(Object(A).name), UCase$(txtSearch))
                If B > 0 Then lstList.AddItem CStr(A) + ": " + Object(A).name
            Else
                lstList.AddItem CStr(A) + ": " + Object(A).name
            End If
        Next A
    Case modeMonsters
        For A = 1 To MaxTotalMonsters
            If Len(txtSearch) > 0 Then
                B = InStr(UCase$(Monster(A).name), UCase$(txtSearch))
                If B > 0 Then lstList.AddItem CStr(A) + ": " + Monster(A).name
            Else
                lstList.AddItem CStr(A) + ": " + Monster(A).name
            End If
        Next A
    Case modeNPCs
        For A = 1 To MaxNPCs
            If Len(txtSearch) > 0 Then
                B = InStr(UCase$(NPC(A).name), UCase$(txtSearch))
                If B > 0 Then lstList.AddItem CStr(A) + ": " + NPC(A).name
            Else
                lstList.AddItem CStr(A) + ": " + NPC(A).name
            End If
        Next A
    Case modeHalls
        For A = 1 To MaxHalls
            If Len(txtSearch) > 0 Then
                B = InStr(UCase$(Hall(A).name), UCase$(txtSearch))
                If B > 0 Then lstList.AddItem CStr(A) + ": " + Hall(A).name
            Else
                lstList.AddItem CStr(A) + ": " + Hall(A).name
            End If
        Next A
    Case modeMagic
        For A = 1 To MaxMagic
            If Len(txtSearch) > 0 Then
                B = InStr(UCase$(Magic(A).name), UCase$(txtSearch))
                If B > 0 Then lstList.AddItem CStr(A) + ": " + Magic(A).name
            Else
                lstList.AddItem CStr(A) + ": " + Magic(A).name
            End If
        Next A
    Case modeBans

    Case modePrefix
        For A = 1 To MaxModifications
            If Len(txtSearch) > 0 Then
                B = InStr(UCase$(ItemPrefix(A).name), UCase$(txtSearch))
                If B > 0 Then lstList.AddItem CStr(A) + ": " + ItemPrefix(A).name
            Else
                lstList.AddItem CStr(A) + ": " + ItemPrefix(A).name
            End If
        Next A
    Case modeSuffix
        For A = 1 To MaxModifications
            If Len(txtSearch) > 0 Then
                B = InStr(UCase$(ItemSuffix(A).name), UCase$(txtSearch))
                If B > 0 Then lstList.AddItem CStr(A) + ": " + ItemSuffix(A).name
            Else
                lstList.AddItem CStr(A) + ": " + ItemSuffix(A).name
            End If
        Next A
    End Select
End Sub

Private Sub Form_Load()
    txtSearch.Text = ""
    DrawList
End Sub

Private Sub lstList_DblClick()
    If btnOk.Enabled = True Then
        btnOk_Click
    End If
End Sub

Private Sub txtSearch_Change()
    DrawList
End Sub
