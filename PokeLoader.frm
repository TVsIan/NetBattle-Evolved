VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form PokeLoader 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin NetBattle.CompressZIt CompressZIt1 
      Left            =   600
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin CCRProgressBar6.ccrpProgressBar Progress 
      Height          =   375
      Left            =   120
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer AnimTimer 
      Interval        =   250
      Left            =   1320
      Top             =   1800
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeLoader.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeLoader.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeLoader.frx":11B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Pokémon Data..."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1920
      Top             =   1080
      Width           =   480
   End
End
Attribute VB_Name = "PokeLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AnimFrame As Byte
Dim Done As Boolean
Const DoDBExport As Boolean = False

Private Sub AnimTimer_Timer()
    AnimFrame = AnimFrame + 1
    If AnimFrame = 5 Then AnimFrame = 1
    Select Case AnimFrame
        Case 1
            Image1.Picture = ImageList1.ListImages(1).Picture
        Case 2, 4
            Image1.Picture = ImageList1.ListImages(2).Picture
        Case 3
            Image1.Picture = ImageList1.ListImages(3).Picture
    End Select
    If Done Then
        '>>> Call WriteDebugLog("Unloading PokeLoader.")
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Done = False
    DoEvents
    Me.Left = (Screen.Width - Me.Width) \ 2
    If SplashScreenUp Then
        Me.Top = frmSplash.Top + frmSplash.Height
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
    End If
    Me.Show
    AnimFrame = 1
    Image1.Picture = ImageList1.ListImages(1).Picture
    '>>> Call WriteDebugLog("Data load started")
    Call LoadPKMNData
End Sub

Public Sub LoadPKMNData()
    'Load everything out of the database.
    'I'm not commenting everything because it should be fairly self-explanatory.
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Temp As String
    Dim TempVar As String
    Dim MTTemp As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim CurrentRecord As Integer
    Dim TempPKMN As Pokemon
    Dim TempPDex As PokeDexInfo
    Dim TempMove As Move
    Dim TempRow(17) As Single
    Dim CompressedDB() As Byte
    Dim FileSize As Long
    Dim RawMoves As String
    Dim RawMachine As String
    Dim RawBreeding As String
    Dim RawRBY As String
    Dim RawRBYTM As String
    Dim RawSpecial As String
    Dim RawTutor As String
    Dim RawAdv As String
    Dim RawAdvTM As String
    Dim RawAdvBreed As String
    Dim RawAdvSpecial As String
    Dim RawAdvTutor As String
    Dim RawLFOnly As String
    Dim Offset As Byte
    Dim FileNum As Integer
    Dim CSVFile As String
    Dim TempArray() As String
    Dim RawMoveLevels(0 To 3) As String
    Dim MoveTemp As Integer
    Dim LevTemp As Byte
    Dim DBE As Long
    Dim TempEx() As String
    Dim TempRaw(1 To 13) As String
    ReDim BasePKMN(0) As Pokemon
    ReDim PDexText(0) As PokeDexInfo
    
    'Dunno where else to stick this, so...
    For X = 0 To 24
        Y = X \ 5 + 1
        Nature(X).StatChg(Y) = Nature(X).StatChg(Y) + 1
        Y = X Mod 5 + 1
        Nature(X).StatChg(Y) = Nature(X).StatChg(Y) - 1
    Next X
    
    FileNum = FreeFile
    BasePKMN(0).Name = "???"
    Progress.Value = 0
    Progress.Max = 760 '386
    '>>> Call WriteDebugLog("PokeDB Complete - Load MoveDB...")
    Open SlashPath & "MoveDB.cdb" For Input As #FileNum
    Input #FileNum, FileSize
    Close #FileNum
    ReDim CompressedDB(FileLen(SlashPath & "MoveDB.cdb") - (Len(Str(FileSize)) - 2)) As Byte
    Open SlashPath & "MoveDB.cdb" For Binary Access Read As #FileNum
    Get #FileNum, Len(Str(FileSize)) + 2, CompressedDB()
    Close #FileNum
    CompressZIt1.DecompressData CompressedDB(), FileSize
    CSVFile = "MoveDB" & FixedHex(Int(Rnd * 65536), 4) & ".csv"
    Open SlashPath & CSVFile For Binary Access Write As #FileNum
    Put #FileNum, , CompressedDB()
    Close #FileNum
    Open SlashPath & CSVFile For Input As #FileNum
    CurrentRecord = 0
    'Progress.Value = 0
    'Progress.Max = 354
    Seek #FileNum, 1
    While Not EOF(FileNum)
        With TempMove
            Input #FileNum, .ID, .Name, .Type, .Power, .Accuracy, .PP, .SpecialPercent, .SpecialEffect, .Target, .Text, .WorksRight, .BrightPowder, .KingsRock, .RBYMove, .GSCMove, .AdvMove, .HitsTeam, .SelfMove, .OldTM, .NewTM, .ADVTM, .SubstituteBlocks, .HitsAll, .SoundMove, .PhysMove, .MagicCoat
            If .ID <> CurrentRecord + 1 Then
                Debug.Print "Screws up after Move # " & CurrentRecord
            End If
            CurrentRecord = .ID
        End With
        If UBound(Moves) < CurrentRecord Then
            ReDim Preserve Moves(CurrentRecord) As Move
        End If
        Moves(CurrentRecord) = TempMove
        Label1.Caption = "Loading Moves: " & Moves(CurrentRecord).Name
        Progress.Value = Progress.Value + 1
        If Progress.Value Mod 10 = 0 Then DoEvents
    Wend
    Close #FileNum
    Kill SlashPath & CSVFile
    MainContainer.MousePointer = vbHourglass
    '>>> Call WriteDebugLog("Load PokeDB...")
    Open SlashPath & "PokeDB.cdb" For Input As #FileNum
    Input #FileNum, FileSize
    Close #FileNum
    ReDim CompressedDB(FileLen(SlashPath & "PokeDB.cdb") - Len(Str(FileSize)) - 2) As Byte
    Open SlashPath & "PokeDB.cdb" For Binary Access Read As #FileNum
    Get #FileNum, Len(Str(FileSize)) + 2, CompressedDB()
    Close #FileNum
    CompressZIt1.DecompressData CompressedDB(), FileSize
    CSVFile = "PokeDB" & FixedHex(Int(Rnd * 65536), 4) & ".csv"
    Open SlashPath & CSVFile For Binary Access Write As #FileNum
    Put #FileNum, , CompressedDB()
    Close #FileNum
    CurrentRecord = 0
    Open SlashPath & CSVFile For Input As #FileNum
    Seek #FileNum, 1
    
    If DoDBExport Then
        DBE = FreeFile
        Open SlashPath & "DBExport.csv" For Output As #DBE
        Write #DBE, "**DBExport**", "RBY Level", "RBY TM", "GSC Level", "GSC TM", "GSC Egg", "GSC Special", "Crystal Tutor", "RS Level", "RS TM", "RS Egg", "RS Special", "LF Tutor", "LF Level"
    End If
    
    While Not EOF(FileNum)
        With TempPKMN
            Input #FileNum, .No, .GSNo, .AdvNo, .Name, .Legendary, .Uber, .Type1, .Type2, .PAtt(0), .PAtt(1), .Color1, .Color2, .BaseHP, .BaseAttack, .BaseDefense, .BaseSpeed, .BaseSAttack, .BaseSDefense, .BaseSpecial, .StartsWith, RawMoves, RawMachine, RawBreeding, RawRBY, RawRBYTM, RawSpecial, RawTutor, RawAdv, RawAdvTM, RawAdvBreed, RawAdvSpecial, RawAdvTutor, RawLFOnly, .ExistRBY, .ExistGSC, .ExistAdv, .PercentFemale, TempPDex.RedBlue, TempPDex.Yellow, TempPDex.Gold, TempPDex.Silver, TempPDex.Crystal, TempPDex.Ruby, TempPDex.Sapphire, .MyStage, .MyMethod, .Evo(1), .EvoM(1), .Evo(2), .EvoM(2), .Evo(3), .EvoM(3), .Evo(4), .EvoM(4), .Evo(5), .EvoM(5), .Weight, .Height, .Offset, .LevelBal, .EggGroup1, .EggGroup2, .Illegals(0), .Illegals(1), .Illegals(2), .Illegals(3), .BreedIllegals(0), .BreedIllegals(1), .BreedIllegals(2), .BreedIllegals(3), RawMoveLevels(0), RawMoveLevels(1), RawMoveLevels(2), RawMoveLevels(3)
            If .No <> CurrentRecord + 1 Then
                Debug.Print "Screws up after " & CurrentRecord
            End If
            CurrentRecord = .No
            
            If DoDBExport Then
                TempRaw(1) = RawRBY
                TempRaw(2) = RawRBYTM
                TempRaw(3) = RawMoves
                TempRaw(4) = RawMachine
                TempRaw(5) = RawBreeding
                TempRaw(6) = RawSpecial
                TempRaw(7) = ""
                TempRaw(8) = RawAdv
                TempRaw(9) = RawAdvTM
                TempRaw(10) = RawAdvBreed
                TempRaw(11) = RawAdvSpecial
                TempRaw(12) = RawAdvTutor
                TempRaw(13) = RawLFOnly
                If Len(RawTutor) > 0 Then
                    MTTemp = RawTutor
                    If MTTemp - 4 >= 0 Then
                        TempRaw(7) = TempRaw(7) & "70,"
                        MTTemp = MTTemp - 4
                    End If
                    If MTTemp - 2 >= 0 Then
                        TempRaw(7) = TempRaw(7) & "98,"
                        MTTemp = MTTemp - 2
                    End If
                    If MTTemp - 1 >= 0 Then
                        TempRaw(7) = TempRaw(7) & "232,"
                        MTTemp = MTTemp - 1
                    End If
                End If

                For Y = 1 To 13
                    If Right$(TempRaw(Y), 1) = "," Then
                        TempRaw(Y) = Left$(TempRaw(Y), Len(TempRaw(Y)) - 1)
                    End If
                    TempEx = Split(TempRaw(Y), ",")
                    Select Case Y
                    Case 2
                        For Z = 0 To UBound(TempEx)
                            TempEx(Z) = Moves(Val(TempEx(Z))).OldTM & Moves(Val(TempEx(Z))).Name
                        Next Z
                        SortStringArray TempEx
                        For Z = 0 To UBound(TempEx)
                            ChopString TempEx(Z), 4
                        Next Z
                    Case 4
                        For Z = 0 To UBound(TempEx)
                            TempEx(Z) = Moves(Val(TempEx(Z))).NewTM & Moves(Val(TempEx(Z))).Name
                        Next Z
                        SortStringArray TempEx
                        For Z = 0 To UBound(TempEx)
                            ChopString TempEx(Z), 4
                        Next Z
                    Case 9
                        For Z = 0 To UBound(TempEx)
                            TempEx(Z) = Moves(Val(TempEx(Z))).ADVTM & Moves(Val(TempEx(Z))).Name
                        Next Z
                        SortStringArray TempEx
                        For Z = 0 To UBound(TempEx)
                            ChopString TempEx(Z), 4
                        Next Z
                    Case Else
                        For Z = 0 To UBound(TempEx)
                            TempEx(Z) = Moves(Val(TempEx(Z))).Name
                        Next Z
                        SortStringArray TempEx
                    End Select
                    TempRaw(Y) = Join(TempEx, ", ")
                Next Y
                Write #DBE, .Name, TempRaw(1), TempRaw(2), TempRaw(3), TempRaw(4), TempRaw(5), TempRaw(6), TempRaw(7), TempRaw(8), TempRaw(9), TempRaw(10), TempRaw(11), TempRaw(12), TempRaw(13)
            End If
            
        End With
        If UBound(BasePKMN) < CurrentRecord Then
            ReDim Preserve BasePKMN(CurrentRecord) As Pokemon
            ReDim Preserve PDexText(CurrentRecord) As PokeDexInfo
        End If
        BasePKMN(CurrentRecord) = TempPKMN
        PDexText(CurrentRecord) = TempPDex
        
        With BasePKMN(CurrentRecord)
            ReDim .RBYMoves(1)
            ReDim .RBYTM(1)
            ReDim .BaseMoves(1)
            ReDim .MachineMoves(1)
            ReDim .BreedingMoves(1)
            ReDim .SpecialMoves(1)
            ReDim .MoveTutor(3)
            ReDim .AdvMoves(1)
            ReDim .ADVTM(1)
            ReDim .AdvBreeding(1)
            ReDim .AdvSpecial(1)
            ReDim .AdvTutor(1)
            ReDim .LFOnly(1)
            If Len(RawMoves) > 0 Then
                TempArray = Split(RawMoves, ",")
                ReDim .BaseMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .BaseMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawMachine) > 0 Then
                TempArray = Split(RawMachine, ",")
                ReDim .MachineMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .MachineMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawBreeding) > 0 Then
                TempArray = Split(RawBreeding, ",")
                ReDim .BreedingMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .BreedingMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawRBY) > 0 Then
                TempArray = Split(RawRBY, ",")
                ReDim .RBYMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .RBYMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawRBYTM) > 0 Then
                TempArray = Split(RawRBYTM, ",")
                ReDim .RBYTM(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .RBYTM(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawSpecial) > 0 Then
                TempArray = Split(RawSpecial, ",")
                ReDim .SpecialMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .SpecialMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawTutor) > 0 Then
                MTTemp = RawTutor
                If MTTemp - 4 >= 0 Then
                    .MoveTutor(1) = 70
                    MTTemp = MTTemp - 4
                End If
                If MTTemp - 2 >= 0 Then
                    .MoveTutor(2) = 98
                    MTTemp = MTTemp - 2
                End If
                If MTTemp - 1 >= 0 Then
                    .MoveTutor(3) = 232
                    MTTemp = MTTemp - 1
                End If
            End If
            If Len(RawAdv) > 0 Then
                TempArray = Split(RawAdv, ",")
                ReDim .AdvMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvMoves(Y) = Val(TempArray(Y - 1))
                Next Y
                .TotalAdvMoves = UBound(TempArray)
            Else
                .TotalAdvMoves = 0
            End If
            If Len(RawAdvTM) > 0 Then
                TempArray = Split(RawAdvTM, ",")
                ReDim .ADVTM(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .ADVTM(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawAdvBreed) > 0 Then
                TempArray = Split(RawAdvBreed, ",")
                ReDim .AdvBreeding(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvBreeding(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawAdvSpecial) > 0 Then
                TempArray = Split(RawAdvSpecial, ",")
                ReDim .AdvSpecial(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvSpecial(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawAdvTutor) > 0 Then
                TempArray = Split(RawAdvTutor, ",")
                ReDim .AdvTutor(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvTutor(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(RawLFOnly) > 0 Then
                TempArray = Split(RawLFOnly, ",")
                ReDim .LFOnly(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .LFOnly(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            'Fill in L100 stats for the Pokedex
            .Attack = GetStat(100, .BaseAttack, 15)
            .Defense = GetStat(100, .BaseDefense, 15)
            .Speed = GetStat(100, .BaseSpeed, 15)
            .SpecialAttack = GetStat(100, .BaseSAttack, 15)
            .SpecialDefense = GetStat(100, .BaseSDefense, 15)
            .MaxHP = GetHP(100, .BaseHP, 15)
            ReDim .MoveLevel(UBound(Moves), 3) As Byte
            'If RawMoveLevels(2) <> "" Then Debug.Print .No, RawMoveLevels(2)
            For X = 0 To 3
                Y = 0
                While Y < Len(RawMoveLevels(X))
                    MoveTemp = Val(Mid(RawMoveLevels(X), Y + 1, 3))
                    LevTemp = Val(Mid(RawMoveLevels(X), Y + 4, 3))
                    .MoveLevel(MoveTemp, X) = LevTemp
                    Y = Y + 6
                Wend
            Next
            
            .ModAttr(0) = .PAtt(0)
            .ModAttr(1) = .PAtt(1)
        End With
        Label1.Caption = "Loading Pokémon Data: " & BasePKMN(CurrentRecord).Name
        Progress.Value = Progress.Value + 1
        If Progress.Value Mod 10 = 0 Then DoEvents
    Wend
    If DoDBExport Then Close #DBE
    Close #FileNum
    Kill SlashPath & CSVFile
    '>>> Call WriteDebugLog("MoveDB Complete - Load TypeDB...")
    Open SlashPath & "TypeDB.cdb" For Input As #FileNum
    Input #FileNum, FileSize
    Close #FileNum
    ReDim CompressedDB(FileLen(SlashPath & "TypeDB.cdb") - (Len(Str(FileSize)) - 2)) As Byte
    Open SlashPath & "TypeDB.cdb" For Binary Access Read As #FileNum
    Get #FileNum, Len(Str(FileSize)) + 2, CompressedDB()
    Close #FileNum
    CompressZIt1.DecompressData CompressedDB(), FileSize
    CSVFile = "TypeDB" & FixedHex(Int(Rnd * 65536), 4) & ".csv"
    Open SlashPath & CSVFile For Binary Access Write As #FileNum
    Put #FileNum, , CompressedDB()
    Close #FileNum
    Open SlashPath & CSVFile For Input As #FileNum
    'Progress.Value = 0
    'Progress.Max = 17
    CurrentRecord = 0
    While Not EOF(FileNum)
        Input #FileNum, CurrentRecord, TempRow(1), TempRow(2), TempRow(3), TempRow(4), TempRow(5), TempRow(6), TempRow(7), TempRow(8), TempRow(9), TempRow(10), TempRow(11), TempRow(12), TempRow(13), TempRow(14), TempRow(15), TempRow(16), TempRow(17)
        For X = 1 To 17
            BattleMatrix(CurrentRecord, X) = TempRow(X)
        Next
        Label1.Caption = "Loading Type Data: " & Element(CurrentRecord)
        Progress.Value = Progress.Value + 1
        If Progress.Value Mod 10 = 0 Then DoEvents
    Wend
    Close #FileNum
    Kill SlashPath & CSVFile
    MainContainer.MousePointer = vbNormal
    Done = True
    '>>> Call WriteDebugLog("Data load complete - waiting for close.")
End Sub
