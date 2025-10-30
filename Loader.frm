VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Loader 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickStart"
   ClientHeight    =   2550
   ClientLeft      =   240
   ClientTop       =   1260
   ClientWidth     =   3975
   ClipControls    =   0   'False
   HelpContextID   =   20001
   Icon            =   "Loader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Loader.frx":1272
   ScaleHeight     =   2550
   ScaleWidth      =   3975
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2295
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6959
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   6000
      Pattern         =   "*.pnl"
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   600
      Left            =   2040
      Top             =   1560
      WhatsThisHelpID =   10007
      Width           =   1875
   End
   Begin VB.Image Button4Up 
      Height          =   585
      Left            =   7200
      Picture         =   "Loader.frx":4D61
      Top             =   3960
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Button4Down 
      Height          =   585
      Left            =   9120
      Picture         =   "Loader.frx":585F
      Top             =   3960
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   120
      Top             =   1560
      WhatsThisHelpID =   10006
      Width           =   1875
   End
   Begin VB.Image Button3Down 
      Height          =   600
      Left            =   9120
      Picture         =   "Loader.frx":630D
      Top             =   3360
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Button3Up 
      Height          =   600
      Left            =   7200
      Picture         =   "Loader.frx":6E83
      Top             =   3360
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   600
      Left            =   120
      Top             =   840
      WhatsThisHelpID =   10005
      Width           =   3750
   End
   Begin VB.Image Button2Disable 
      Height          =   600
      Left            =   7200
      Picture         =   "Loader.frx":7A72
      Top             =   2640
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image Button2Down 
      Height          =   600
      Left            =   7200
      Picture         =   "Loader.frx":8B49
      Top             =   2040
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image Button2Up 
      Height          =   600
      Left            =   7200
      Picture         =   "Loader.frx":9B07
      Top             =   1440
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image Button1Down 
      Height          =   600
      Left            =   7200
      Picture         =   "Loader.frx":ACFB
      Top             =   720
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image Button1Up 
      Height          =   600
      Left            =   7200
      Picture         =   "Loader.frx":BC49
      Top             =   120
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   120
      Top             =   120
      WhatsThisHelpID =   10004
      Width           =   3750
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New..."
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&1 (No Recent File)"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&2 (No Recent File)"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&3 (No Recent File)"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&4 (No Recent File)"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   8
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuBattle 
      Caption         =   "&Battle"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuBattleItem 
         Caption         =   "&Start Battling"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "&Advanced..."
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "&Replay..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuPokedex 
      Caption         =   "&DataDex"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&PokéDex"
         Index           =   0
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&MoveDex"
         Index           =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&BattleDex"
         Index           =   2
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&DamageCalc"
         Index           =   3
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Set Options..."
         Index           =   0
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Wizard..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "D&ebug"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Battle Window..."
         Index           =   0
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Master Server..."
         Index           =   1
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Play Battle Music (Test)"
         Index           =   2
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Dump Ranks"
         Index           =   3
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Reload Database"
         Index           =   4
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Edit Pokémon..."
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "S&tadium Mode Window..."
         Index           =   6
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "Edit Battle &Chart..."
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Pokédex Parser..."
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&View Station ID..."
         Index           =   9
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Template"
         Index           =   10
         Begin VB.Menu mnuTemplateItem 
            Caption         =   "Dump &Move Template"
            Index           =   0
         End
         Begin VB.Menu mnuTemplateItem 
            Caption         =   "Dump &Pokemon Template"
            Index           =   1
         End
         Begin VB.Menu mnuTemplateItem 
            Caption         =   "Dump &Flavor Text Template"
            Index           =   2
         End
         Begin VB.Menu mnuTemplateItem 
            Caption         =   "Dump Poke&Dex Template"
            Index           =   3
         End
         Begin VB.Menu mnuTemplateItem 
            Caption         =   "Dump Mis&cellaneous Template"
            Index           =   4
         End
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Imagelist Dump"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Rearranger..."
         Index           =   12
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Advanced Arranger..."
         Index           =   13
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Item Switcher..."
         Index           =   14
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Extract Unowns & Substitute"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Compress CSV Databases"
         Index           =   16
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Export Movelist"
         Index           =   17
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Help..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Web Site..."
         Index           =   1
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About"
         Index           =   3
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Loader.frm
'This is the initial window.
'I called it Loader since it loads up the database stuff.
'It's been through several revisions, and there's probably some junk I can get rid of.
Option Explicit
Private BWorked As Boolean
'Private Type Movesets
'    mSet() As String
'    UB As Integer
'End Type
'
'Private Sub Command5_Click()
'    Dim Temp As String
'    Dim State As Integer
'    Dim CurrentMove As Integer
'    Dim M As Integer
'    Dim Poke(1 To 386, 1 To 3) As Movesets
'    Dim X As Integer
'    Dim Y As Integer
'    State = 1
'    Open "C:\RS386Alpha.txt" For Input As #1
'    Do Until EOF(1)
'        Line Input #1, Temp
'        If Temp = "" Then
'            State = 1
'        Else
'            If State = 1 Then
'                Select Case Temp
'                Case "Level Up"
'                    M = 1
'                Case "TM/HM"
'                    M = 2
'                Case "Breeding"
'                    M = 3
'                Case Else
'                    CurrentMove = GetMoveNum(Temp)
'                End Select
'                State = 0
'                Line Input #1, Temp
'            Else
'                If CurrentMove = 0 Then Stop
'                X = InStr(1, Temp, " at")
'                If X <> 0 Then Temp = Left(Temp, X - 1)
'                X = InStr(1, Temp, " from")
'                If X <> 0 Then Temp = Left(Temp, X - 1)
'                X = InStr(1, Temp, " as")
'                If X <> 0 Then Temp = Left(Temp, X - 1)
'                Y = GetPokeNum(Temp)
'                For X = 0 To Poke(Y, M).UB - 1
'                    If Val(Poke(Y, M).mSet(X)) >= CurrentMove Then Exit For
'                Next X
'                If X < Poke(Y, M).UB Then
'                    If Val(Poke(Y, M).mSet(X)) = CurrentMove Then X = -1
'                End If
'                If X >= 0 Then
'                    ReDim Preserve Poke(Y, M).mSet(Poke(Y, M).UB)
'                    For X = Poke(Y, M).UB - 1 To X Step -1
'                        Poke(Y, M).mSet(X + 1) = Poke(Y, M).mSet(X)
'                    Next X
'                    Poke(Y, M).mSet(X + 1) = CStr(CurrentMove)
'                    Poke(Y, M).UB = Poke(Y, M).UB + 1
'                End If
'            End If
'        End If
'    Loop
'    Close #1
'    Open "C:\moveoutput.txt" For Output As #1
'    For X = 1 To 386
'        Print #1, Join(Poke(X, 1).mSet, ",") & ", " & Join(Poke(X, 2).mSet, ",") & ", " & Join(Poke(X, 3).mSet, ",") & ", "
'    Next X
'    Close #1
'End Sub


'All these Command#_Click items are left over from the early versions.
'It had regular buttons instead of the image buttons (note to self: make a custom control out of that).
'They get called up by the image's MouseUp event.
Private Sub Command1_Click()
    'Open Team Builder.
    On Error Resume Next
    If BasePKMN(1).No <> 1 Then PokeLoader.Show
    Loader.Visible = False
    Unload TeamBuilder
    Call WriteDebugLog("TB Click")
    TeamBuilder.Show
End Sub

Private Sub Command2_Click()
    'Battle
    ServerAddress = ""
    If Not BattleEligible(PKMN()) Then Exit Sub
    Call DoBattle
    'If BWorked Then Me.Visible = False
End Sub

Sub DoBattle()
    Me.Visible = False
    DoEvents
    BWorked = False
    Select Case ServerAddress
    Case "", "GoTo Listing"
        ServerList.Show 1
        Call DoBattle
    Case "Error"
        NetSet.Show 1
        Call DoBattle
    Case "Cancelled"
        Me.Visible = True
    Case Else
        BWorked = True
        DoEvents
        MasterServer.Show
        'Unload Me
    End Select
'    Select Case GameType
'        Case 0
'            If ServerAddress = "" Then ServerList.Show 1
'            If ServerAddress = "Error" Then
'                BWorked = False
'                Exit Sub
'            Else
'                MasterServer.Show
'                BWorked = True
'            End If
'        Case 1
'            MasterServer.Show
'            BWorked = True
'        Case 2
'            MsgBox "This function isn't available yet.", vbCritical, "Error"
'            BWorked = False
'        Case 3
'            MsgBox "This function isn't available yet.", vbCritical, "Error"
'            BWorked = False
'        Case 4
'            MsgBox "This function isn't available yet.", vbCritical, "Error"
'            BWorked = False
'    End Select
End Sub

'Command3 was Exit.

Private Sub Command4_Click()
    'About box
    frmAbout.Show 1
End Sub


Private Sub Command6_Click()
    Dim Temp As String
    Dim Temp2 As String
    Dim X As Integer
    Dim Y As Integer
'    Open "D:\Sapp Dex.txt" For Input As #1
'    Open "D:\Sapp Dex 2.txt" For Output As #2
    Open "D:\Ruby Dex.txt" For Input As #1
    Open "D:\Ruby Dex 2.txt" For Output As #2
    X = -1
    Do Until EOF(1)
        X = X + 1
        Line Input #1, Temp
        Temp = Replace(Temp, "POKéMON", "Pokémon")
        Temp = Replace(Temp, "TRAINER", "Trainer")
        Temp = Replace(Temp, "POKé BALL", "Poké Ball")
        Temp = Replace(Temp, "NIDORAN", "Nidoran")
        For Y = 1 To 17
            Temp = Replace(Temp, Element(Y), Element(Y), , , vbTextCompare)
        Next Y
        For Y = 1 To 386
            Temp = Replace(Temp, BasePKMN(Y).Name, BasePKMN(Y).Name, , , vbTextCompare)
        Next Y
        For Y = 1 To Len(Temp) - 1
            If LCase(Mid(Temp, Y, 1)) <> Mid(Temp, Y, 1) And LCase(Mid(Temp, Y + 1, 1)) <> Mid(Temp, Y + 1, 1) Then
                Temp2 = Temp2 & X & " "
                Exit For
            End If
        Next Y
        Print #2, Temp
    Loop
    Close #1
    Close #2
    MsgBox Temp2
End Sub

Private Sub Command5_Click()
    Dim X As Integer
    Dim Y As Long
    Dim Temp As String
    Dim Temp2 As String
    Dim Build As String
    
    For X = 1 To UBound(BasePKMN)
        Build = Build & BasePKMN(X).Name & vbNullChar
    Next X
    For X = 1 To UBound(Moves)
        Build = Build & Moves(X).Name & vbNullChar
    Next X
    Temp = Join(Item, vbNullChar)
    If Left$(Temp, 1) = vbNullChar Then ChopString Temp, 1
    Build = Build & Temp & vbNullChar
    Temp = Join(AttributeText, vbNullChar)
    If Left$(Temp, 1) = vbNullChar Then ChopString Temp, 1
    Build = Build & Temp & vbNullChar
    Temp = Join(Element, vbNullChar)
    If Left$(Temp, 1) = vbNullChar Then ChopString Temp, 1
    Build = Build & Temp & vbNullChar
    For X = 0 To UBound(Nature)
        Build = Build & Nature(X).Name & vbNullChar
    Next X
    Temp = Join(StatName, vbNullChar)
    If Left$(Temp, 1) = vbNullChar Then ChopString Temp, 1
    Build = Build & Temp
    
    Y = Len(Build)
    MainContainer.Compressor.CompressString Build
    X = Len(Build)
    Temp = "    "
    CopyMemory ByVal Temp, Y, ByVal 4
    Temp2 = Temp
    Temp = "  "
    CopyMemory ByVal Temp, X, ByVal 2
    Build = Temp2 & Temp & Build
    Temp = "  "
    X = UBound(FTextLen)
    CopyMemory ByVal Temp, X, ByVal 2
    Build = Build & Temp
    Temp = String$(X, " ")
    For X = 1 To X
        Mid(Temp, X, 1) = Chr$(FTextLen(X))
    Next X
    Build = Build & Temp
    Temp = ""
    For X = 1 To X - 1
        Temp = Temp & GetFText(X)
    Next X
    MainContainer.Compressor.CompressString Temp
    Build = Build & Temp
    Open SlashPath & "English.pnf" For Binary Access Write As #1
    Put #1, , Build
    Close #1
End Sub

Private Sub Command11_Click()
    Dim Build As String
    Dim X As Long
    For X = 1 To UBound(Moves)
        Build = Build & "Moves(" & X & ") = " & Chr(34) & LCase(Replace(Moves(X).Name, " ", "")) & Chr(34) & vbNewLine
    Next X
    Clipboard.Clear
    Clipboard.SetText Build
End Sub

Private Sub Command12_Click()
    Open "C:\DBOutput.csv" For Output As #1
    Dim X As Long
    Write #1, "No.", "Name", "HP", "Atk", "Def", "Spd", "SAtk", "SDef"
    For X = 1 To 389
        With BasePKMN(X)
            Write #1, .No, .Name, GetAdvHP(.BaseHP, 31, 255, 100), GetAdvStat(.BaseAttack, 31, 255, 100, 0), GetAdvStat(.BaseDefense, 31, 255, 100, 0), GetAdvStat(.BaseSpeed, 31, 255, 100, 0), GetAdvStat(.BaseSAttack, 31, 255, 100, 0), GetAdvStat(.BaseSDefense, 31, 255, 100, 0)
        End With
    Next X
    Close #1
End Sub

Private Sub Command100_Click()
    Dim Poke As Long
    Dim Temp As String
    Dim M As Long
    Dim X As Long
    Dim U As Long
    Open "C:\emerald.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, Temp
        Temp = Trim$(Temp)
        If Len(Temp) = 0 Then
            Poke = 0
        ElseIf Poke = 0 Then
            Poke = GetPokeNum(Temp)
            If Poke = 0 Then Stop
        Else
            M = GetMoveNum(Temp)
            If M = 0 Then Stop
            With BasePKMN(Poke)
                U = UBound(.AdvTutor)
                For X = 1 To U
                    If .AdvTutor(X) = M Then
                        Exit For
                    End If
                Next X
                If X > U Then
                    ReDim Preserve .AdvTutor(X)
                    .AdvTutor(X) = M
                End If
            End With
        End If
    Loop
    Close #1
    Open "C:\output.txt" For Output As #1
    For X = 1 To 389
        Temp = ""
        For M = 1 To UBound(BasePKMN(X).AdvTutor)
            Temp = Temp & CStr(BasePKMN(X).AdvTutor(M)) & ","
        Next M
        If Temp <> "0," Then Print #1, Temp Else Print #1, ""
    Next X
    Close #1
    
End Sub

Private Sub Command672_Click()
    BenchDLL
End Sub

'Private Sub Command7_Click()
'    Dim Build As String
'    Dim Temp As String
'    Dim T() As Integer
'    Dim T2() As Integer
'    Dim T3() As Integer
'    'Dim A() As Integer
'    Dim A() As String
'    Dim X As Long
'    Dim Y As Long
'    Dim Z As Long
'    Dim B As Long
'    For X = 1 To 389
'        Build = Build & "*****" & UCase(BasePKMN(X).Name) & "*****" & vbNewLine
'        Call MakeMoveArray(X, nbRBYTrade, T, A)
'        Call MakeMoveArray(X, nbGSCTrade, T2, A)
'        Call MakeMoveArray(X, nbFullAdvance, T3, A)
'        Erase A
'        Z = UBound(T)
'        ReDim Preserve T(1 To Z + UBound(T2) + UBound(T3))
'        B = 0
'        For Y = Z + 1 To UBound(T2) + Z
'            B = B + 1
'            T(Y) = T2(B)
'        Next Y
'        B = 0
'        For Y = Z + UBound(T2) + 1 To UBound(T3) + UBound(T2) + Z
'            B = B + 1
'            T(Y) = T3(B)
'        Next Y
'
'        For Y = 1 To UBound(T)
'            For Z = Y + 1 To UBound(T)
'                If T(Y) = T(Z) Then T(Z) = 0
'            Next Z
'        Next Y
'
'        For Y = 1 To UBound(T)
'            If T(Y) > 0 Then
'                If (Not A) = -1 Then ReDim A(0) Else ReDim Preserve A(UBound(A) + 1)
'                A(UBound(A)) = Moves(T(Y)).Name
'            End If
'        Next Y
'        If (Not A) <> -1 Then
'            Call SortStringArray(A)
'            Build = Build & Join(A, ", ") & vbNewLine
'        End If
'
'
''        Call MakeMoveArray(X, nbTrueRBY, T, A)
''        Erase A
''        For Y = 1 To UBound(T)
''            If T(Y) > 0 Then
''                If (Not A) = -1 Then ReDim A(0) Else ReDim Preserve A(UBound(A) + 1)
''                A(UBound(A)) = Moves(T(Y)).Name
''            End If
''        Next Y
''        If (Not A) <> -1 Then
''            Call SortStringArray(A)
''            Build = Build & "True RBY:  " & Join(A, ", ") & vbNewLine
''        End If
''
''        Call MakeMoveArray(X, nbRBYTrade, T, A)
''        Erase A
''        For Y = 1 To UBound(T)
''            If T(Y) > 0 Then
''                If (Not A) = -1 Then ReDim A(0) Else ReDim Preserve A(UBound(A) + 1)
''                A(UBound(A)) = Moves(T(Y)).Name
''            End If
''        Next Y
''        If (Not A) <> -1 Then
''            Call SortStringArray(A)
''            Build = Build & "RBY w/ Trades:  " & Join(A, ", ") & vbNewLine
''        End If
''
''        Call MakeMoveArray(X, nbTrueGSC, T, A)
''        Erase A
''        For Y = 1 To UBound(T)
''            If T(Y) > 0 Then
''                If (Not A) = -1 Then ReDim A(0) Else ReDim Preserve A(UBound(A) + 1)
''                A(UBound(A)) = Moves(T(Y)).Name
''            End If
''        Next Y
''        If (Not A) <> -1 Then
''            Call SortStringArray(A)
''            Build = Build & "True GSC:  " & Join(A, ", ") & vbNewLine
''        End If
''
''        Call MakeMoveArray(X, nbGSCTrade, T, A)
''        Erase A
''        For Y = 1 To UBound(T)
''            If T(Y) > 0 Then
''                If (Not A) = -1 Then ReDim A(0) Else ReDim Preserve A(UBound(A) + 1)
''                A(UBound(A)) = Moves(T(Y)).Name
''            End If
''        Next Y
''        If (Not A) <> -1 Then
''            Call SortStringArray(A)
''            Build = Build & "GSC w/ Trades:  " & Join(A, ", ") & vbNewLine
''        End If
''
''        Call MakeMoveArray(X, nbTrueRuSa, T, A)
''        Erase A
''        For Y = 1 To UBound(T)
''            If T(Y) > 0 Then
''                If (Not A) = -1 Then ReDim A(0) Else ReDim Preserve A(UBound(A) + 1)
''                A(UBound(A)) = Moves(T(Y)).Name
''            End If
''        Next Y
''        If (Not A) <> -1 Then
''            Call SortStringArray(A)
''            Build = Build & "True RuSa:  " & Join(A, ", ") & vbNewLine
''        End If
''
''        Call MakeMoveArray(X, nbFullAdvance, T, A)
''        Erase A
''        For Y = 1 To UBound(T)
''            If T(Y) > 0 Then
''                If (Not A) = -1 Then ReDim A(0) Else ReDim Preserve A(UBound(A) + 1)
''                A(UBound(A)) = Moves(T(Y)).Name
''            End If
''        Next Y
''        If (Not A) <> -1 Then
''            Call SortStringArray(A)
''            Build = Build & "Full RuSa:  " & Join(A, ", ") & vbNewLine
''        End If
'    Next X
'    Open "C:\ftproot\DBOutput.txt" For Binary Access Write As #1
'    Put #1, , Build
'    Close #1
'End Sub

'(De)Activate Debug mode if capital D is pressed. (REMOVED)
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 68 Then
'        DebugMode = Not DebugMode
'        mnuDebug.Visible = DebugMode
'    End If
'End Sub

Private Sub Form_Load()
    Dim FileToUse As String
    Dim TempVar As String
    Dim X As Integer
    Dim DidMessage As Boolean
    Dim OldMajor As Integer
    Dim OldMinor As Integer
    Dim OldRevision As Integer
    Dim RevisionRange As Integer
    Dim WindowTitle As String
    Dim WindowText As String
    Dim Answer As Integer
    Dim CommandLine As String
    On Error Resume Next
    WriteDebugLog "Loader Form_Load executing"
    If InVBMode Then mnuDebug.Visible = True
    If MainContainer.LoaderKiller.Enabled Then Exit Sub
    'MsgBox ">" & Command$ & "<"
    CommandLine = Command$
    
    DoneLoading = False
    'Hide itself
    Me.Hide
    '>>> Call WriteDebugLog("Centering Loader...")
    Call CenterWindow(Loader)
    
    DoEvents
    If MainContainer.LoaderKiller.Enabled Then Exit Sub
    
    'Set the initial state on the buttons
    '>>> Call WriteDebugLog("Setting pictures...")
    WriteDebugLog "Setting pictures"
    Image1.Picture = Button1Up.Picture
    Image2.Picture = Button2Disable.Picture
    Image3.Picture = Button3Up.Picture
    Image4.Picture = Button4Up.Picture
    
    '>>> Call WriteDebugLog("Setting recent files...")
    Call RefreshRecentFiles
    
    'Get the list of language files
    '>>> Call WriteDebugLog("Getting available languages...")
    LFile(0).Text = "None (US English)"
    LFile(0).FileName = ""
    LFile(0).HasBattle = True
    LFile(0).HasMisc = True
    LFile(0).HasMoves = True
    LFile(0).HasPDEX = True
    LFile(0).HasPKMN = True
    LFile(0).HasProgram = True
    File1.Path = SlashPath
    Call LoadLang
    
    WriteDebugLog "Loading settings"
    'Load the settings
    'If in Safe mode, use the defaults.
    If UCase(CommandLine) = "/SAFE" Then
        '>>> Call WriteDebugLog("Safe Mode enabled")
        SoundOption = 0
        MusicOption = 0
        AnimOption = 0
        FancyText = True
        AllowViewing = 1
        LogPrompt = 1
        AutoScan = 0
        AskOnUpdate = 1
        BMessStyle = 1
        CurrLang = ""
        GetSpeed = 1
        TBSort = 0
        TBMode = 2
        Autosave = 0
        AddLineBreaks = False
        UseBG = False
        LogSave = 0
        ReplayPrompt = 0
        OldInterface = False
        UseTS = False
        TSFormat = "[hh:mm:ss]"
        ParseURLs = False
        VerIcons = False
        ColorNames = False
        Autoload = False
        DisplayLines = 1000
        CSFilter = Split("", ";")
        'UseDX = False
    Else
        '>>> Call WriteDebugLog("Loading settings...")
        SoundOption = GetSetting("NetBattle", "Options", "ServerSound", 0)
        MusicOption = GetSetting("NetBattle", "Options", "Music", 0)
        AnimOption = GetSetting("NetBattle", "Options", "Animation", 0)
        'FancyText = GetSetting("NetBattle", "Options", "Fancy Text", True)
        FancyText = True
        AllowViewing = GetSetting("NetBattle", "Options", "Allow Viewing", 1)
        Autoload = GetSetting("NetBattle", "Options", "Autoload", True)
        LogPrompt = GetSetting("NetBattle", "Options", "Log Prompt", 1)
        AutoScan = GetSetting("NetBattle", "Options", "Auto Scan", -1)
        AskOnUpdate = GetSetting("NetBattle", "Options", "Ask On Update", 1)
        BMessStyle = GetSetting("NetBattle", "Options", "Message Style", 1)
        CurrLang = GetSetting("NetBattle", "Options", "Language", "")
        GetSpeed = GetSetting("NetBattle", "Options", "GetSpeed", 1)
        DoMultiPaste = GetSetting("NetBattle", "Options", "DoMultiPaste", True)
        TBSort = GetSetting("NetBattle", "Options", "Team Builder Sort", 0)
        If TBSort = 2 Then TBSort = 0
        Autosave = GetSetting("NetBattle", "Options", "Save Replays", 0)
        AddLineBreaks = GetSetting("NetBattle", "Options", "Line Breaks", False)
        UseBG = GetSetting("NetBattle", "Options", "Use Background", True)
        LogSave = GetSetting("NetBattle", "Options", "Log Save", 0)
        ReplayPrompt = GetSetting("NetBattle", "Options", "Replay Prompt", 0)
        TBMode = GetSetting("NetBattle", "Options", "LastTBMode", 3)
        If TBMode < 0 Or TBMode > 6 Or TBMode = 4 Then TBMode = 2
        OldInterface = GetSetting("NetBattle", "Options", "Old Interface", False)
        MoveDelay = GetSetting("NetBattle", "Options", "MoveDelay", 1500)
        UseNicks = GetSetting("NetBattle", "Options", "UseNicks", True)
        UsePrefix = GetSetting("NetBattle", "Options", "UsePrefix", False)
        If MoveDelay > 5000 Then MoveDelay = 5000
        If MoveDelay < 0 Then MoveDelay = 0
        UseTS = GetSetting("NetBattle", "Options", "UseTS", False)
        TSFormat = GetSetting("NetBattle", "Options", "TSFormat", "[hh:mm:ss]")
        ParseURLs = GetSetting("NetBattle", "Options", "ParseURLs", True)
        DisplayLines = GetSetting("NetBattle", "Options", "Lines", 1000)
        VerIcons = GetSetting("NetBattle", "Options", "VerIcons", False)
        ColorNames = GetSetting("NetBattle", "Options", "ColorNames", False)
        TempVar = Dec2Bin(GetSetting("NetBattle", "Options", "MsgToggles", 1023), 10)
        CSFilter = Split(GetSetting("NetBattle", "Options", "CSFilter", vbNullString), ";")
        For X = 1 To 10
            MsgToggle(X) = CBool(Mid$(TempVar, X, 1))
        Next X
        
        'UseDX = GetSetting("NetBattle", "DirectX", "Use", True)
        UseHiResTimer = GetSetting("NetBattle", "DirectX", "Timer", True)
        RefreshRate = GetSetting("NetBattle", "DirectX", "Interval", 2)
        DeviceGUID = GetSetting("NetBattle", "DirectX", "Device", "")
    End If
    '>>> Call WriteDebugLog("Getting sound files")
    For X = 0 To 10
        SoundFile(X) = GetSetting("NetBattle", "Sound", Trim(Str(X)), "")
        If SoundFile(X) = "" Then Call ResetDefaultSound(X)
        If Not FileExists(SoundFile(X)) Then Call ResetDefaultSound(X)
        SoundEnable(X) = GetSetting("NetBattle", "Enable Sound", Trim(Str(X)), True)
        If Not FileExists(SoundFile(X)) Then SoundEnable(X) = False
    Next
    'Loaded settings that aren't affected by Safe Mode.
    '>>> Call WriteDebugLog("Getting personalized settings")
    SavedPassword = GetSetting("NetBattle", "Options", "Saved Password", "")
    If Len(SavedPassword) <> 32 And Len(SavedPassword) > 0 Then SavedPassword = MD5(SavedPassword)
    SaveSetting "NetBattle", "Options", "Saved Password", SavedPassword
    DidMessage = GetSetting("NetBattle", "Notification", "DidPasswordMessage", False)
    OldMajor = GetSetting("NetBattle", "Notification", "OldMajor", -1)
    OldMinor = GetSetting("NetBattle", "Notification", "OldMinor", -1)
    OldRevision = GetSetting("NetBattle", "Notification", "OldRevision", -1)
    'Run the Wizard on the first load
    '>>> Call WriteDebugLog("Loading CfgWiz")
    If AutoScan = -1 Then CfgWiz.Show 1
    AutoScan = Abs(AutoScan)
    
    'What's New if this is a different version than the last run.
    'Keeping the last 5 revisions in here.
    '>>> Call WriteDebugLog("Checking Update Text. " & App.Major & "." & App.Minor & "." & App.Revision & " -- " & OldMajor & "." & OldMinor & "." & OldRevision)
    If App.Major <> OldMajor Or App.Minor <> OldMinor Or App.Revision <> OldRevision Then
        If OldRevision = -1 Then WindowTitle = "Welcome!" Else WindowTitle = "Upgrade Detected"
        If OldMajor = App.Major And OldMinor = App.Minor Or OldMajor = -1 Then
            WindowText = "Recent Changes:" & vbNewLine
            If OldRevision > App.Revision Then OldRevision = -1
            If App.Revision - OldRevision > 5 Then RevisionRange = App.Revision - 5 Else RevisionRange = OldRevision + 1
            For X = RevisionRange To App.Revision
                Select Case X
                Case 6
                    WindowText = WindowText & vbNewLine & "0.9.6: General bugfixes."
                Case 5
                    WindowText = WindowText & vbNewLine & "0.9.5: " & vbNewLine & " Battle resume" & vbNewLine & " Custom database mods" & vbNewLine & " Undo Move" & vbNewLine & " Damage Calculator" & vbNewLine & " Various bugfixes" & vbNewLine & " See website for full changelist"
                Case 4
                    WindowText = WindowText & vbNewLine & "0.9.4: Emerald and bugfix release."
                Case 3
                    WindowText = WindowText & vbNewLine & "0.9.3: Advance release, major changes.  See website for full details"
                Case 2
                    WindowText = WindowText & vbNewLine & "0.9.2: Emergency fix for Away crashes."
                Case 1
                    WindowText = WindowText & vbNewLine & "0.9.1: Bugfixes to slow connections and a few other things."
                End Select
            Next
        End If
        If WindowText = "" Then WindowText = "0.9.0: Major changes - go see the web site."
        MsgBox WindowText, vbInformation, WindowTitle
        SaveSetting "NetBattle", "Notification", "OldMajor", App.Major
        SaveSetting "NetBattle", "Notification", "OldMinor", App.Minor
        SaveSetting "NetBattle", "Notification", "OldRevision", App.Revision
    End If
    'Auto-Update check
    '>>> Call WriteDebugLog("Doing Version Scan")
    If AutoScan = 1 And Not DidScan Then Call DoVersionScan
    'Recent file list refresh
    Call RefreshRecentFiles
    'Default to NOT on relay server
    RelayServer = False
    'Play the opening music/sound
    '>>> Call WriteDebugLog("Playing Music")
    If MusicOption = 1 Then
        Call PlayMusic(4, False)
    ElseIf SoundOption = 1 Then
        Call PlaySound(0)
    End If
    
    'Load the data if it isn't already
    If BasePKMN(1).No <> 1 Then PokeLoader.Show
    'Load language
    Call DoLanguage
    
    'Re-show itself
    Me.Show
    'Get rid of the splash screen
    Unload frmSplash
    DoneLoading = True
    '>>> Call WriteDebugLog("Done processing startup!")
    
    'Load a team or replay from a windows Explorer double-click
    On Error Resume Next
    If CommandLine <> "" And UCase(CommandLine) <> "/SAFE" Then
        TempVar = Dequote(CommandLine)
        Select Case Right$(TempVar, 4)
        Case ".pnb"
            If CmdTeam Then Exit Sub
            FileToUse = TempVar
            Call TeamLoader.ReadFile(FileToUse)
            Call RefreshBattleButtons
            Call RefreshRecentFiles
            CmdTeam = True
        Case ".btl"
            If CmdReplay Then Exit Sub
            'Fill in replay loading here.
            'Playback.Show
            'Playback.CmdFile = TempVar
            'Call Playback.Toolbar1_ButtonClick(Playback.Toolbar1.Buttons(1))
            CmdReplay = True
        End Select
    ElseIf Autoload And Len(RecentFiles(1)) > 0 Then
        TempVar = RecentFiles(1)
        If FileExists(TempVar) Then
            Call TeamLoader.ReadFile(TempVar)
            Call RefreshBattleButtons
            X = InStrRev(TempVar, "\")
            TempVar = Right$(TempVar, Len(TempVar) - X)
            StatusBar.Panels(1).Text = "Autoloaded " & TempVar
        End If
       
    End If
    
    TempVar = GetSetting("NetBattle", "Recent Files", "Mod", vbNullChar)
    If TempVar <> vbNullChar Then
        TempVar = SlashPath & "Database Mods\" & TempVar & ".mod"
        If FileExists(TempVar) Then LoadDBMod TempVar
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Stop the music on any unload
    If MusicOption = 1 Then Call StopMusic
    'If it wasn't unloaded by code, kill MainContainer (and end the program)
    If UnloadMode = 0 Then
        MainContainer.LoaderKiller.Enabled = True
        Cancel = 1
    End If
End Sub

Private Sub Image1_Click()
    Call Command1_Click
End Sub

'Boring image button code follows.
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image1.Picture = Button1Down.Picture
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image1.Picture = Button1Up.Picture
    End If
End Sub

Private Sub Image2_Click()
    Call Command2_Click
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image2.Picture = Button2Down.Picture
    End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image2.Picture = Button2Up.Picture
    End If
End Sub

Private Sub Image3_Click()
    Call Command4_Click
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image3.Picture = Button3Down.Picture
    End If
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image3.Picture = Button3Up.Picture
    End If
End Sub

Private Sub Image4_Click()
    Unload MainContainer
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image4.Picture = Button4Down.Picture
    End If
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Image4.Picture = Button4Up.Picture
    End If
End Sub


'Menu code follows.
Private Sub mnuBattleItem_Click(Index As Integer)
    Select Case Index
        Case 1
            If Not BattleEligible(PKMN()) Then Exit Sub
            Call Command2_Click
        Case 2
            If Not BattleEligible(PKMN()) Then Exit Sub
            ServerAddress = "Error"
            Call DoBattle
        Case 4
            OpenedAsReplay = True
            Me.Hide
            Battle.Show
    End Select
End Sub

Private Sub mnuDebugItem_Click(Index As Integer)
    Dim X As Integer
    Dim Y As Long
    Dim Answer As String
    Dim TempString As String
    Dim BlankPKMN As Pokemon
    Dim Temp() As String
    
    On Error Resume Next
    Select Case Index
        Case 0
            You.Picture = 6
            Battle.Show
        Case 1
            ServerWindow.Show
        Case 2
            If mnuDebugItem(2).Checked Then
                Call StopMusic
                mnuDebugItem(2).Checked = False
            Else
                Call PlayMusic(5, True)
                mnuDebugItem(2).Checked = True
            End If
        Case 3
            If BasePKMN(1).No <> 1 Then PokeLoader.Show
            Call DoDebugRank
        Case 4
            ReDim BasePKMN(0) As Pokemon
            PokeLoader.Show
            MsgBox "Pokémon data has been cleared and reloaded.", vbInformation, "Done"
        Case 5
        Case 6
            Stadium.Show
        Case 7
        Case 8
            'Call DoPDexEntries
        Case 9
            MsgBox DecompressSID(StationID), vbInformation, "Station ID"
        Case 11
            For X = 1 To MainContainer.Types.ListImages.count
                SavePicture MainContainer.Types.ListImages(X).Picture, SlashPath & X & ".ico"
            Next
        Case 12
            Rearrange.Show
        Case 13
            BoxArrange.Show
        Case 14
            ItemChange.Show
        Case 15
            Call MainContainer.DoPicture("20101.gif", False)
            Call MainContainer.DoPicture("20102.gif", False)
            Call MainContainer.DoPicture("20103.gif", False)
            Call MainContainer.DoPicture("20104.gif", False)
            Call MainContainer.DoPicture("20105.gif", False)
            Call MainContainer.DoPicture("20106.gif", False)
            Call MainContainer.DoPicture("20107.gif", False)
            Call MainContainer.DoPicture("20108.gif", False)
            Call MainContainer.DoPicture("20109.gif", False)
            Call MainContainer.DoPicture("20110.gif", False)
            Call MainContainer.DoPicture("20111.gif", False)
            Call MainContainer.DoPicture("20112.gif", False)
            Call MainContainer.DoPicture("20113.gif", False)
            Call MainContainer.DoPicture("20114.gif", False)
            Call MainContainer.DoPicture("20115.gif", False)
            Call MainContainer.DoPicture("20116.gif", False)
            Call MainContainer.DoPicture("20117.gif", False)
            Call MainContainer.DoPicture("20118.gif", False)
            Call MainContainer.DoPicture("20119.gif", False)
            Call MainContainer.DoPicture("20120.gif", False)
            Call MainContainer.DoPicture("20121.gif", False)
            Call MainContainer.DoPicture("20122.gif", False)
            Call MainContainer.DoPicture("20123.gif", False)
            Call MainContainer.DoPicture("20124.gif", False)
            Call MainContainer.DoPicture("20125.gif", False)
            Call MainContainer.DoPicture("20126.gif", False)
            Call MainContainer.DoPicture("201s.gif", False)
            Call MainContainer.DoPicture("201gb.gif", False)
            Call MainContainer.DoPicture("201gbs.gif", False)
            Call MainContainer.DoPicture("subst.gif", False)
        Case 16

    End Select
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Dim FileToUse As String
    Dim BlankPKMN As Pokemon
    Dim BlankTrainer As Trainer
    Dim X As Integer
    On Error Resume Next
    Select Case Index
        Case 0
            You = BlankTrainer
            If BetaRel <> "" Then
                You.ProgVersion = App.Major & "." & App.Minor & "." & BetaRel
            Else
                You.ProgVersion = App.Major & "." & App.Minor & "." & App.Revision
            End If
            For X = 1 To 6
                PKMN(X) = BlankPKMN
            Next X
            StoredFileName = ""
            Call Command1_Click
        Case 1
            If BasePKMN(1).No <> 1 Then PokeLoader.Show
            Call TeamLoader.OpenTheFile
            Call RefreshBattleButtons
            Call RefreshRecentFiles
        Case 3 To 6
            If BasePKMN(1).No <> 1 Then PokeLoader.Show
            Call TeamLoader.ReadFile(RecentFiles(Index - 2))
            Call RefreshBattleButtons
            Call RefreshRecentFiles
        Case 8
            MainContainer.LoaderKiller.Enabled = True
    End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            Call ShowHelpContext(20001)
        Case 1
            ShellExecute 0, vbNullString, "http://www.netbattle.net", vbNullString, vbNullString, 0
        Case 3
            frmAbout.Show 1
    End Select
End Sub

Private Sub mnuOptionsItem_Click(Index As Integer)
    Select Case Index
        Case 0
            Loader.Visible = False
            Options.Show vbModeless, MainContainer
        Case 1
            CfgWiz.Show 1
    End Select
End Sub

Private Sub mnuPokedexItem_Click(Index As Integer)
    If BasePKMN(1).No <> 1 Then PokeLoader.Show
    MasterDex.Show
    MasterDex.SetMode Index
End Sub

Sub RefreshRecentFiles()
    'Refresh the Recent File listing
    Dim X As Integer
    
    For X = 1 To 4
        If RecentFiles(X) <> "" Then
            mnuFileItem(X + 2).Caption = "&" & X & " " & RecentFiles(X)
            mnuFileItem(X + 2).Enabled = True
        Else
            mnuFileItem(X + 2).Caption = "&" & X & " (No Recent File)"
            mnuFileItem(X + 2).Enabled = False
        End If
    Next X
End Sub

Private Function GoWinInet(sURL$) As String
    'Not my code (except for one small change)
    'Calls up direct downloading from an HTTP server.
    Dim sBuffer As String * 4096
    Dim sReturn As String
    Dim lNumBytes As Long
    Dim lSession As Long
    Dim LFile As Long
    Dim bReadOK As Boolean

    lSession = InternetOpen("NetBattle", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    LFile = InternetOpenUrl(lSession, sURL, vbNullString, 0, INTERNET_FLAG_EXISITING_CONNECT + INTERNET_FLAG_RELOAD, 0)
    If LFile Then
        Do
            bReadOK = InternetReadFile(LFile, sBuffer, Len(sBuffer), lNumBytes)
            If lNumBytes Then
                sReturn = sReturn & Left$(sBuffer, lNumBytes)
            End If
        Loop While bReadOK And lNumBytes > 0
        InternetCloseHandle (LFile)
        GoWinInet = sReturn
    Else
        'MsgBox "Cannot open update page", vbCritical, "Error updating"
        StatusBar.Panels(1).Text = "Cannot access update page"
    End If
End Function

Public Sub DoVersionScan()
    'Check the current version
    'Seems to hang (or pause for a looooooong time) under certain server conditions
    'Safe mode was put in for just such an occasion.
    Dim NewVersion As String
    Dim Answer As Integer
    Dim ProgFile As String
    Dim DLLFile As String
    Dim Started As Double
    Dim InstallerVersion As String
    Dim V1 As Integer
    Dim V2 As Integer
    Dim V3 As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DLFile As Integer
    
    On Error GoTo Failed
    '>>> Call WriteDebugLog("Checking for updates...")
    NewVersion = GoWinInet(BaseURL & "version.txt")
    If Len(Trim(NewVersion)) <> 0 Then
        X = InStr(1, NewVersion, ".")
        Y = InStrRev(NewVersion, ".")
        V1 = Val(Left(NewVersion, X - 1))
        V2 = Val(Mid(NewVersion, X + 1, Y - X - 1))
        V3 = Val(Right(NewVersion, Len(NewVersion) - Y))
    Else
        StatusBar.Panels(1).Text = "Error accessing update site"
        Exit Sub
    End If
    'Version.txt contains only a string
    'with the current version (ie "0.8.36").
    'This one covers the server being down.
    If InStr(1, NewVersion, "<HTML>") > 0 Or Len(Trim(NewVersion)) = 0 Then
        StatusBar.Panels(1).Text = "Unable to connect to update server."
    ElseIf IsVersionAt(You.ProgVersion, V1, V2, V3) Or (BetaRel <> "" And IsVersionAt(App.Major & "." & App.Minor & "." & Left(BetaRel, Len(BetaRel) - 1), V1, V2, V3)) Then
        StatusBar.Panels(1).Text = "You are using the most current version."
    'Hey, new version available!
    'This will also come up if the server has an OLDER version,
    'But that shouldn't happen except on development systems.
    
    'Not anymore... I finally got sick of it.  ~Masa
    Else
        StatusBar.Panels(1).Text = "Version " & NewVersion & " is available for download."
        'Prompt if set to ask, otherwise just go ahead.
        If AskOnUpdate = 1 Then
            Answer = MsgBox("You are currently using version " & You.ProgVersion & vbNewLine & NewVersion & " is available.  Would you like to automatically update?", vbQuestion + vbYesNo + vbDefaultButton1, "Update Available")
        Else
            Answer = vbYes
        End If
        If Answer = vbYes Then
            Loader.Visible = True
            InstallerVersion = GoWinInet(BaseURL & "version2.txt")
            'Version2.txt is the update installer's current version.
            'Needed to do it that way to overwrite Pokebattle.exe
            'This bit download the two needed files (if they're not already there),
            'and registers the OCX.
            DLFile = FreeFile
            'NOT DOWNLOADING THESE ANYMORE!  I'll just use a Web installer, much easier.
            'If Not FileExists(SysDir & "CompressZIt.ocx") Then
            '    StatusBar.Panels(1).Text = "Downloading installer resources..."
            '    NewVersion = GoWinInet(BaseURL & "CompressZIt.ocx")
            '    Open SysDir & "CompressZIt.ocx" For Output As #DLFile
            '        Print #DLFile, NewVersion
            '    Close
            '    Shell SysDir & "Regsvr32.exe /s " & SysDir & "CompressZIt.ocx"
            'End If
            'If Not FileExists(SysDir & "zlib.dll") Then
            '    StatusBar.Panels(1).Text = "Downloading installer resources..."
            '    NewVersion = GoWinInet(BaseURL & "zlib.dll")
            '    Open SysDir & "zlib.dll" For Output As #DLFile
            '        Print #DLFile, NewVersion
            '    Close
            'End If
            'This downloads the installer itself.
            StatusBar.Panels(1).Text = "Downloading installer..."
            NewVersion = GoWinInet(BaseURL & "instupdt.exe." & InstallerVersion)
            If Right(App.Path, 1) = "\" Then ProgFile = App.Path & "instupdt.exe" Else ProgFile = App.Path & "\instupdt.exe"
            Open ProgFile For Output As #DLFile
                Print #DLFile, NewVersion
            Close #DLFile
            'Start the installer and close yourself.
            Started = Shell(ProgFile, vbNormalFocus)
            End
            Exit Sub
        End If
    End If
    'Don't scan again this session.
    DidScan = True
    Exit Sub
Failed:
    '>>> Call WriteDebugLog("Update failed.")
    'MsgBox "Automatic update failed.  You may not be connected to the Internet, or the server may be down.", vbCritical, "Error"
    StatusBar.Panels(1).Text = "Automatic Update failed - See web site"
End Sub

Function PrepareString(ByVal Original As String) As String
    Dim Temp As Long
    Dim Temp2 As Long
    Dim NewString As String
    
    NewString = Original
    Temp = 1
    While InStr(Temp, NewString, "'") > 0
        Temp2 = InStr(Temp, NewString, "'")
        NewString = Left(NewString, Temp2 - 1) & "'" & Right(NewString, Len(NewString) - Temp2 + 1)
        Temp2 = InStr(Temp, NewString, "'")
        Temp = Temp2 + 2
    Wend
    PrepareString = NewString
End Function

Private Sub mnuTemplateItem_Click(Index As Integer)
    Dim X As Integer
    Dim FileNum As Integer
        
    FileNum = FreeFile
    Select Case Index
        Case 0
            If Moves(1).ID = 0 Then
                MsgBox "Database not loaded!", vbCritical, "Error"
                Exit Sub
            End If
            Open SlashPath & "template.pnm" For Output As #FileNum
            For X = 1 To UBound(Moves)
                Write #FileNum, Moves(X).ID, Moves(X).Name, Moves(X).Text
            Next
            Close #FileNum
            MsgBox "Move template dumped to template.pnm", vbInformation, "Done"
        Case 1
            If BasePKMN(1).Name = "" Then
                MsgBox "Database not loaded!", vbCritical, "Error"
                Exit Sub
            End If
            Open SlashPath & "template.pnp" For Output As #FileNum
            For X = 1 To UBound(BasePKMN)
                Write #FileNum, BasePKMN(X).No, BasePKMN(X).Name
            Next
            Close #FileNum
            MsgBox "Pokemon template dumped to template.pnp", vbInformation, "Done"
        Case 2
'            Open SlashPath & "template.pnf" For Output As #FileNum
'            For X = 1 To UBound(FlavorText)
'                Write #FileNum, X, FlavorText(X)
'            Next
'            Close #FileNum
'            MsgBox "Flavor Text template dumped to template.pnf", vbInformation, "Done"
        Case 3
            Open SlashPath & "template.pnd" For Output As #FileNum
            For X = 1 To UBound(PDexText)
                With PDexText(X)
                    Write #FileNum, X, .RedBlue, .Yellow, .Gold, .Silver, .Crystal, .Ruby, .Sapphire
                End With
            Next
            Close #FileNum
            MsgBox "Pokedex template dumped to template.pnd", vbInformation, "Done"
        Case 4
            Open SlashPath & "template.pnr" For Output As #FileNum
            For X = 1 To UBound(Element)
                Write #1, "E", X, Element(X)
            Next
            For X = 0 To UBound(Gender)
                Write #1, "G", X, Gender(X)
            Next
            For X = 1 To UBound(Condition)
                Write #1, "C", X, Condition(X)
            Next
            For X = 0 To UBound(Item)
                Write #1, "I", X, Item(X)
            Next
            For X = 0 To UBound(Weather)
                Write #1, "W", X, Weather(X)
            Next
            For X = 1 To UBound(RuleText)
                Write #1, "R", X, RuleText(X)
            Next
            For X = 1 To UBound(RuleToolTip)
                Write #1, "T", X, RuleToolTip(X)
            Next
            For X = 0 To UBound(EvoMethod)
                Write #1, "M", X, EvoMethod(X)
            Next
            For X = 0 To UBound(ColorText)
                Write #1, "L", X, ColorText(X)
            Next
            For X = 0 To UBound(AttributeText)
                Write #1, "A", X, AttributeText(X)
            Next
            For X = 1 To UBound(StatName)
                Write #1, "S", X, StatName(X)
            Next
            For X = 1 To UBound(ModeText)
                Write #1, "D", X, ModeText(X)
            Next
            For X = 0 To UBound(TerrainText)
                Write #1, "N", X, TerrainText(X)
            Next
            Close #FileNum
            MsgBox "Miscellaneous template dumped to template.pnr", vbInformation, "Done"
    End Select
End Sub

Public Sub DoLanguage()
    Dim X As Integer
    Dim Temp As String
    Dim Temp2 As String
    Dim Temp3 As String
    Dim Temp4 As String
    Dim Temp5 As String
    Dim Temp6 As String
    Dim Temp7 As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    'Reset to defaults if no language
    If CurrLang = "" And DoneLoading = True Then
        Call SetEnglish
        PokeLoader.Show
        Exit Sub
    End If
        
    '>>> Call WriteDebugLog("Setting language to " & CurrLang & "...")
    'Apply move names & descriptions
    If FileExists(SlashPath & CurrLang & ".pnm") Then
        Open SlashPath & CurrLang & ".pnm" For Input As #FileNum
        While Not EOF(FileNum)
            Input #FileNum, X, Temp, Temp2
            Moves(X).Name = Temp
            Moves(X).Text = Temp2
        Wend
        Close #FileNum
    End If
    
    'Apply Pokemon names
    If FileExists(SlashPath & CurrLang & ".pnp") Then
        Open SlashPath & CurrLang & ".pnp" For Input As #FileNum
        While Not EOF(FileNum)
            Input #FileNum, X, Temp
            BasePKMN(X).Name = Temp
        Wend
        Close #FileNum
    End If
    
   'Apply flavor text
'    If FileExists(SlashPath & CurrLang & ".pnf") Then
'        Open SlashPath & CurrLang & ".pnf" For Input As #FileNum
'        While Not EOF(FileNum)
'            Input #FileNum, X, Temp
'            FlavorText(X) = Temp
'        Wend
'        Close #FileNum
'    End If
    
    'Apply Pokedex text
    If FileExists(SlashPath & CurrLang & ".pnd") Then
        Open SlashPath & CurrLang & ".pnd" For Input As #FileNum
        While Not EOF(FileNum)
            Input #FileNum, X, Temp, Temp2, Temp3, Temp4, Temp5, Temp6, Temp7
            With PDexText(X)
                .RedBlue = Temp
                .Yellow = Temp2
                .Gold = Temp3
                .Silver = Temp4
                .Crystal = Temp5
                .Ruby = Temp6
                .Sapphire = Temp7
            End With
        Wend
        Close #FileNum
    End If
    
    'Apply Miscellaneous Text
    If FileExists(SlashPath & CurrLang & ".pnr") Then
        Open SlashPath & CurrLang & ".pnr" For Input As #FileNum
        While Not EOF(FileNum)
            Input #FileNum, Temp2, Temp, Temp3
            Select Case Temp2
                Case "E"
                    Element(Temp) = Temp3
                Case "G"
                    Gender(Temp) = Temp3
                Case "C"
                    Condition(Temp) = Temp3
                Case "I"
                    Item(Temp) = Temp3
                Case "W"
                    Weather(Temp) = Temp3
                Case "R"
                    RuleText(Temp) = Temp3
                Case "T"
                    RuleToolTip(Temp) = Temp3
                Case "M"
                    EvoMethod(Temp) = Temp3
                Case "L"
                    ColorText(Temp) = Temp3
                Case "A"
                    AttributeText(Temp) = Temp3
                Case "S"
                    StatName(Temp) = Temp3
                Case "D"
                    ModeText(Temp) = Temp3
                Case "N"
                    TerrainText(Temp) = Temp3
            End Select
        Wend
        Close #FileNum
    End If
    '>>> Call WriteDebugLog("Language change complete.")
End Sub

Sub LoadLang()
    Dim X As Integer
    Dim FileNum As Integer
    
    On Error Resume Next
    If File1.ListCount = 0 Then Exit Sub
    FileNum = FreeFile
    For X = 0 To File1.ListCount - 1
        Open SlashPath & File1.List(X) For Input As #FileNum
        Input #FileNum, LFile(X + 1).Text
        Close #FileNum
        '.pnl = Pokemon Netbattle Language
        LFile(X + 1).FileName = Left(File1.List(X), Len(File1.List(X)) - 4)
        '.pnm = Pokemon Netbattle Moves
        LFile(X + 1).HasMoves = FileExists(SlashPath & LFile(X + 1).FileName & ".pnm")
        '.pnp = Pokemon Netbattle Pokemon
        LFile(X + 1).HasPKMN = FileExists(SlashPath & LFile(X + 1).FileName & ".pnp")
        '.pnd = Pokemon Netbattle pokeDex
        LFile(X + 1).HasPDEX = FileExists(SlashPath & LFile(X + 1).FileName & ".pnd")
        '.pnf = Pokemon Netbattle Flavor text (battle messages)
        LFile(X + 1).HasBattle = FileExists(SlashPath & LFile(X + 1).FileName & ".pnf")
        '.pnt = Pokemon Netbattle Translation (menus, buttons, dialogs)
        LFile(X + 1).HasProgram = FileExists(SlashPath & LFile(X + 1).FileName & ".pnt")
        '.pne = Pokemon Netbattle Random text (elements, items, weather, etc.)
        LFile(X + 1).HasMisc = FileExists(SlashPath & LFile(X + 1).FileName & ".pnr")
    Next
End Sub



Public Sub RefreshBattleButtons()
    Dim Temp As String
    On Error GoTo UnloadCatcher
    If MainContainer.LoaderKiller.Enabled Then Exit Sub
    Temp = BattleOK
    If Temp = "" Then
        Image2.Enabled = True
        Image2.Picture = Button2Up.Picture
        StatusBar.Panels(1).Text = "Ready to battle!"
        mnuBattleItem(1).Enabled = True
        mnuBattleItem(2).Enabled = True
    Else
        Image2.Enabled = False
        Image2.Picture = Button2Disable.Picture
        StatusBar.Panels(1).Text = BattleOK
        mnuBattleItem(1).Enabled = False
        mnuBattleItem(2).Enabled = False
    End If
UnloadCatcher:
End Sub
