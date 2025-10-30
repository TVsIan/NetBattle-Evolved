VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DBCompile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DB Compiler"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin DBMan.CompressZIt CompressZIt1 
      Left            =   120
      Top             =   780
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ProgressBar StepProgBar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1980
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblStep 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "DBCompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Public Sub DoCompile()
    Dim QueryResults As New ADODB.Recordset
    Dim BlankMove As Move
    Dim TempMove As Move
    Dim RowID As Integer
    Dim RowData(17) As Single
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Integer
    Dim M1 As Integer
    Dim M2 As Integer
    Dim M3 As Integer
    Dim M4 As Integer
    Dim TempArray() As String
    Dim OtherArray() As String
    Dim MTTemp As String
    Dim TempMoves() As Integer
    Dim TempSet() As Integer
    Dim Temp As String
    Dim Build As String
    Dim B As Boolean
    Dim TempPKMN As Pokemon
    Dim T As Single
    
    T = Timer
    On Error Resume Next
    ReDim BasePKMN(1 To 389)
    lblStep.Caption = "Converting Move Database"
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Moves WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    Open SlashPath & "MoveDB.csv" For Output As #1
    StepProgBar.Value = 0
    StepProgBar.Max = 354
    While Not QueryResults.EOF
        TempMove = BlankMove
        With TempMove
            .ID = QueryResults("ID")
            .Name = QueryResults("Name")
            .Type = QueryResults("Type")
            .Power = QueryResults("Power")
            .Accuracy = QueryResults("Accuracy")
            .PP = QueryResults("PP")
            .SpecialPercent = QueryResults("Percent")
            .SpecialEffect = QueryResults("Special")
            .Target = QueryResults("Target")
            .Text = QueryResults("Description")
            .WorksRight = QueryResults("Works Properly")
            .BrightPowder = QueryResults("BrightPowder")
            .KingsRock = QueryResults("KingsRock")
            .RBYMove = QueryResults("RBYCompatible")
            .GSCMove = QueryResults("GSCCompatible")
            .AdvMove = QueryResults("ADVCompatible")
            .HitsTeam = QueryResults("HitsBoth")
            .SelfMove = QueryResults("AffectsSelf")
            .OldTM = QueryResults("RBYTM")
            .NewTM = QueryResults("GSTM")
            .ADVTM = QueryResults("AdvTM")
            .HitsAll = QueryResults("HitsAll")
            .SoundMove = QueryResults("SoundMove")
            .PhysMove = QueryResults("PhysMove")
            .SubstituteBlocks = QueryResults("BlockSubstitute")
            .Text = Replace(.Text, Chr(34), "''")
            .MagicCoat = QueryResults("MagicCoat")
            Write #1, .ID, .Name, .Type, .Power, .Accuracy, .PP, .SpecialPercent, .SpecialEffect, .Target, .Text, .WorksRight, .BrightPowder, .KingsRock, .RBYMove, .GSCMove, .AdvMove, .HitsTeam, .SelfMove, .OldTM, .NewTM, .ADVTM, .SubstituteBlocks, .HitsAll, .SoundMove, .PhysMove, .MagicCoat
        End With
        QueryResults.MoveNext
        StepProgBar.Value = StepProgBar.Value + 1
        DoEvents
    Wend
    
    lblStep.Caption = "Compressing Move Database"
    Close #1
    ReDim CBytes(FileLen(SlashPath & "MoveDB.csv") - 1) As Byte
    Open SlashPath & "MoveDB.csv" For Binary Access Read As #1
    Get #1, , CBytes()
    Close #1
    CompressZIt1.CompressData CBytes()
    If FileExists(SlashPath & "header.tmp") Then Kill SlashPath & "header.tmp"
    Open SlashPath & "header.tmp" For Output As #1
    Write #1, CompressZIt1.OriginalSize
    Close #1
    ReDim HBytes(FileLen(SlashPath & "header.tmp") - 1) As Byte
    Open SlashPath & "header.tmp" For Binary Access Read As #1
    Get #1, , HBytes()
    Close #1
    Kill SlashPath & "header.tmp"
    If FileExists(SlashPath & "MoveDB.cdb") Then Kill SlashPath & "MoveDB.cdb"
    Open SlashPath & "MoveDB.cdb" For Binary Access Write As #1
    Put #1, , HBytes
    Put #1, , CBytes
    Close #1
    
    lblStep.Caption = "Converting Type Chart"
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM BattleChart WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    StepProgBar.Value = 0
    StepProgBar.Max = 17
    Open SlashPath & "TypeDB.csv" For Output As #1
    While Not QueryResults.EOF
        RowID = 0
        For X = 1 To 17
            RowData(X) = 0
        Next
        RowID = QueryResults("ID")
        For X = 1 To 17
            RowData(X) = QueryResults(X)
        Next
        Write #1, RowID, RowData(1), RowData(2), RowData(3), RowData(4), RowData(5), RowData(6), RowData(7), RowData(8), RowData(9), RowData(10), RowData(11), RowData(12), RowData(13), RowData(14), RowData(15), RowData(16), RowData(17)
        QueryResults.MoveNext
        StepProgBar.Value = StepProgBar.Value + 1
        DoEvents
    Wend
    Close

    lblStep.Caption = "Compressing Type Chart"
    ReDim CBytes(FileLen(SlashPath & "TypeDB.csv") - 1) As Byte
    Open SlashPath & "TypeDB.csv" For Binary Access Read As #1
    Get #1, , CBytes()
    Close
    CompressZIt1.CompressData CBytes()
    If FileExists(SlashPath & "header.tmp") Then Kill SlashPath & "header.tmp"
    Open SlashPath & "header.tmp" For Output As #1
    Write #1, CompressZIt1.OriginalSize
    Close
    ReDim HBytes(FileLen(SlashPath & "header.tmp") - 1) As Byte
    Open SlashPath & "header.tmp" For Binary Access Read As #1
    Get #1, , HBytes()
    Close
    Kill SlashPath & "header.tmp"
    If FileExists(SlashPath & "TypeDB.cdb") Then Kill SlashPath & "TypeDB.cdb"
    Open SlashPath & "TypeDB.cdb" For Binary Access Write As #1
    Put #1, , HBytes
    Put #1, , CBytes
    Close
    
    lblStep.Caption = "Reading Pokémon Database"
    On Error Resume Next
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Pokemon WHERE Number > 0 ORDER BY Number ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    StepProgBar.Value = 0
    StepProgBar.Max = 389
    While Not QueryResults.EOF
        With BasePKMN(StepProgBar.Value + 1)
            .No = QueryResults("Number")
            .GSNo = QueryResults("GSCNumber")
            .AdvNo = QueryResults("AdvNumber")
            .Name = QueryResults("Name")
            .Legendary = QueryResults("Legendary")
            .Uber = QueryResults("Uber")
            .Type1 = QueryResults("Type1")
            .Type2 = QueryResults("Type2")
            .PAtt(0) = QueryResults("Attribute1")
            .PAtt(1) = QueryResults("Attribute2")
            .Color1 = QueryResults("Color1")
            .Color2 = QueryResults("Color2")
            .BaseHP = QueryResults("HP")
            .BaseAttack = QueryResults("Attack")
            .BaseDefense = QueryResults("Defense")
            .BaseSpeed = QueryResults("Speed")
            .BaseSAttack = QueryResults("SpecialAttack")
            .BaseSDefense = QueryResults("SpecialDefense")
            .BaseSpecial = QueryResults("SpecialRBY")
            .StartsWith = QueryResults("BornWith")
            .RawMoves = QueryResults("Moves")
            .RawMachine = QueryResults("Machine Moves")
            .RawBreeding = QueryResults("Breeding Moves")
            .RawRBY = QueryResults("R/B/Y Moves")
            .RawRBYTM = QueryResults("R/B/Y Machines")
            .RawSpecial = QueryResults("Special Moves")
            .RawTutor = QueryResults("Move Tutor")
            .RawAdv = QueryResults("Advance Moves")
            .RawAdvTM = QueryResults("Advance TMs")
            .RawAdvBreed = QueryResults("Advance Breeding")
            .RawAdvSpecial = QueryResults("Advance Special")
            .RawAdvTutor = QueryResults("Advance Tutor")
            .RawLFOnly = QueryResults("LF Only")
            .MoveLevels(0) = QueryResults("RBYLevels")
            .MoveLevels(1) = QueryResults("GSCLevels")
            .MoveLevels(2) = QueryResults("ADVLevels")
            .MoveLevels(3) = QueryResults("LFLevels")
            ReDim .RBYMoves(0)
            ReDim .RBYTM(0)
            ReDim .BaseMoves(0)
            ReDim .MachineMoves(0)
            ReDim .BreedingMoves(0)
            ReDim .SpecialMoves(0)
            ReDim .MoveTutor(3)
            ReDim .AdvMoves(0)
            ReDim .ADVTM(0)
            ReDim .AdvBreeding(0)
            ReDim .AdvSpecial(0)
            ReDim .AdvTutor(0)
            ReDim .LFOnly(0)
            If Len(.RawMoves) > 0 Then
                TempArray = Split(.RawMoves, ",")
                ReDim .BaseMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .BaseMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawMachine) > 0 Then
                TempArray = Split(.RawMachine, ",")
                ReDim .MachineMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .MachineMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawBreeding) > 0 Then
                TempArray = Split(.RawBreeding, ",")
                ReDim .BreedingMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .BreedingMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawRBY) > 0 Then
                TempArray = Split(.RawRBY, ",")
                ReDim .RBYMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .RBYMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawRBYTM) > 0 Then
                TempArray = Split(.RawRBYTM, ",")
                ReDim .RBYTM(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .RBYTM(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawSpecial) > 0 Then
                TempArray = Split(.RawSpecial, ",")
                ReDim .SpecialMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .SpecialMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawTutor) > 0 Then
                MTTemp = .RawTutor
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
            If Len(.RawAdv) > 0 Then
                TempArray = Split(.RawAdv, ",")
                ReDim .AdvMoves(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvMoves(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawAdvTM) > 0 Then
                TempArray = Split(.RawAdvTM, ",")
                ReDim .ADVTM(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .ADVTM(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawAdvBreed) > 0 Then
                TempArray = Split(.RawAdvBreed, ",")
                ReDim .AdvBreeding(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvBreeding(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawAdvSpecial) > 0 Then
                TempArray = Split(.RawAdvSpecial, ",")
                ReDim .AdvSpecial(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvSpecial(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawAdvTutor) > 0 Then
                TempArray = Split(.RawAdvTutor, ",")
                ReDim .AdvTutor(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .AdvTutor(Y) = Val(TempArray(Y - 1))
                Next Y
            End If
            If Len(.RawLFOnly) > 0 Then
                TempArray = Split(.RawLFOnly, ",")
                ReDim .LFOnly(UBound(TempArray))
                For Y = 1 To UBound(TempArray)
                    .LFOnly(Y) = Val(TempArray(Y - 1))
                Next Y
            End If

            .ExistRBY = QueryResults("RBY")
            .ExistGSC = QueryResults("GSC")
            .ExistAdv = QueryResults("Adv")
            .PercentFemale = QueryResults("Percent Female")
            .RedBlue = QueryResults("PokedexRB")
            .Yellow = QueryResults("PokedexYellow")
            .Gold = QueryResults("PokedexGold")
            .Silver = QueryResults("PokedexSilver")
            .Crystal = QueryResults("PokedexCrystal")
            .Ruby = QueryResults("PokedexRuby")
            .Sapphire = QueryResults("PokedexSapphire")
            .MyStage = QueryResults("MyStage")
            .MyMethod = QueryResults("MyMethod")
            .Evo(1) = QueryResults("Evo1")
            .EvoM(1) = QueryResults("EvoM1")
            .Evo(2) = QueryResults("Evo2")
            .EvoM(2) = QueryResults("EvoM2")
            .Evo(3) = QueryResults("Evo3")
            .EvoM(3) = QueryResults("EvoM3")
            .Evo(4) = QueryResults("Evo4")
            .EvoM(4) = QueryResults("EvoM4")
            .Evo(5) = QueryResults("Evo5")
            .EvoM(5) = QueryResults("EvoM5")
            .Weight = QueryResults("Weight")
            .Height = QueryResults("Height")
            .Offset = QueryResults("Offset")
            .LevelBal = QueryResults("LevelBal")
            .RedBlue = Replace(.RedBlue, Chr(34), "''")
            .Yellow = Replace(.Yellow, Chr(34), "''")
            .Gold = Replace(.Gold, Chr(34), "''")
            .Silver = Replace(.Silver, Chr(34), "''")
            .Crystal = Replace(.Crystal, Chr(34), "''")
            .Ruby = Replace(.Ruby, Chr(34), "''")
            .Sapphire = Replace(.Sapphire, Chr(34), "''")
            .Illegals(0) = QueryResults("IllegalFullGSC")
            .Illegals(1) = QueryResults("IllegalTrueGSC")
            .Illegals(2) = QueryResults("IllegalFullAdv")
            .Illegals(3) = QueryResults("IllegalTrueAdv")
            .EggGroup1 = QueryResults("EggGroup1")
            .EggGroup2 = QueryResults("EggGroup2")
        End With
        QueryResults.MoveNext
        StepProgBar.Value = StepProgBar.Value + 1
        DoEvents
    Wend
    On Error GoTo 0
    
    lblStep.Caption = "Scanning for Illegal Breeding Movesets."
    StepProgBar.Value = 0
    StepProgBar.Max = 1556
    'Here we go...
    For Y = 0 To 3
        Select Case Y
        Case 0: BasePKMN(1).GameVersion = nbGSCTrade
        Case 1: BasePKMN(1).GameVersion = nbTrueGSC
        Case 2: BasePKMN(1).GameVersion = nbFullAdvance
        Case 3: BasePKMN(1).GameVersion = nbTrueRuSa
        End Select
        For X = 2 To 389
            BasePKMN(X).GameVersion = BasePKMN(1).GameVersion
            BasePKMN(X).DoneCheck = False
        Next X
        For X = 1 To 389
            CurrentPoke = X
            TempPKMN = BasePKMN(X)
           'If Y = 0 And TempPKMN.Name = "Cleffa" Then Stop
            With TempPKMN
                If Not (.GameVersion = nbTrueGSC And Not .ExistGSC) And Not (.GameVersion = nbTrueRuSa And Not .ExistAdv) Then
                    If Y < 2 Then TempMoves = .BreedingMoves Else TempMoves = .AdvBreeding
                    'If it can bread with Smeargle or if it has no breeding moves, skip it.
                    'Also, only check first evolutions.
                    If Not (EggGroupCheck(X, 235) And Y = 2) And UBound(TempMoves) > 0 And .MyStage = 1 Then
                        Build = Build & UCase(.Name) & vbNewLine
                        .BreedIllegals(Y) = ""
                        ReDim TempSet(1 To 4)
                        Z = UBound(TempMoves)
                      
                        'Singles
                        For M1 = 1 To Z
                            Temp = TempMoves(M1)
                            TempSet(1) = Temp
                            A = UBound(NowChecking) + 1
                            ReDim Preserve NowChecking(A)
                            ReDim Preserve CheckString(A)
                            NowChecking(A) = X
                            CheckString(A) = TempSet(1)
                            B = Not BreedCheck(X, TempSet)
                            Debug.Print "Check Finished, " & IIf(B, "Illegal", "Legal")
                            ReDim NowChecking(0)
                            ReDim CheckString(0)
'                            If KamexCheck(.No, TempSet) = B Then
'                                OtherArray = Split(Temp, "+")
'                                For Z = 0 To UBound(OtherArray)
'                                    OtherArray(Z) = Moves(OtherArray(Z)).Name
'                                Next Z
'                                Temp = Join(OtherArray, " / ")
'                                Build = Build & IIf(B, "Illegal: ", "Legal: ") & vbNewLine & Temp
'                            End If
                            If B Then
                                .BreedIllegals(Y) = .BreedIllegals(Y) & "|" & TempMoves(M1)
                            End If
                        Next M1
                        
                        'Doubles
                        For M1 = 1 To Z
                            For M2 = M1 + 1 To Z
                                Temp = TempMoves(M1) & "+" & TempMoves(M2)
                                TempSet(1) = TempMoves(M1)
                                TempSet(2) = TempMoves(M2)
                                If Not Prevented(TempPKMN, TempSet) Then
                                    A = UBound(NowChecking) + 1
                                    ReDim Preserve NowChecking(A)
                                    ReDim Preserve CheckString(A)
                                    NowChecking(A) = X
                                    CheckString(A) = Temp
                                    Debug.Print "--------Starting Check--------"
                                    B = Not BreedCheck(X, TempSet)
                                    Debug.Print "Check Finished, " & IIf(B, "Illegal", "Legal")
                                    ReDim Preserve NowChecking(0)
                                    ReDim Preserve CheckString(0)
'                                    If KamexCheck(.No, TempSet) = B Then
'                                        OtherArray = Split(Temp, "+")
'                                        For Z = 0 To UBound(OtherArray)
'                                            OtherArray(Z) = Moves(OtherArray(Z)).Name
'                                        Next Z
'                                        Temp = Join(OtherArray, " / ")
'                                        Build = Build & IIf(B, "Illegal: ", "Legal: ") & vbNewLine & Temp
'                                    End If
                                    If B Then
                                        .BreedIllegals(Y) = .BreedIllegals(Y) & "|" & Temp
                                    End If
                                End If
                            Next M2
                        Next M1
                             
                        'Triples
                        For M1 = 1 To Z
                            For M2 = M1 + 1 To Z
                                For M3 = M2 + 1 To Z
                                    Temp = TempMoves(M1) & "+" & TempMoves(M2) & "+" & TempMoves(M3)
                                    TempSet(1) = TempMoves(M1)
                                    TempSet(2) = TempMoves(M2)
                                    TempSet(3) = TempMoves(M3)
                                    If Not Prevented(TempPKMN, TempSet) Then
                                        A = UBound(NowChecking) + 1
                                        ReDim Preserve NowChecking(A)
                                        ReDim Preserve CheckString(A)
                                        NowChecking(A) = X
                                        CheckString(A) = Temp
                                        B = Not BreedCheck(X, TempSet)
                                        Debug.Print "Check Finished, " & IIf(B, "Illegal", "Legal")
                                        ReDim Preserve NowChecking(0)
                                        ReDim Preserve CheckString(0)
'                                        If KamexCheck(.No, TempSet) = B Then
'                                            OtherArray = Split(Temp, "+")
'                                            For Z = 0 To UBound(OtherArray)
'                                                OtherArray(Z) = Moves(OtherArray(Z)).Name
'                                            Next Z
'                                            Temp = Join(OtherArray, " / ")
'                                            Build = Build & IIf(B, "Illegal: ", "Legal: ") & vbNewLine & Temp
'                                        End If
                                        If B Then
                                            .BreedIllegals(Y) = .BreedIllegals(Y) & "|" & Temp
                                        End If
                                    End If
                                Next M3
                            Next M2
                        Next M1
            
                        'Quads
                        For M1 = 1 To Z
                            For M2 = M1 + 1 To Z
                                For M3 = M2 + 1 To Z
                                    For M4 = M3 + 1 To Z
                                        Temp = TempMoves(M1) & "+" & TempMoves(M2) & "+" & TempMoves(M3) & "+" & TempMoves(M4)
                                        TempSet(1) = TempMoves(M1)
                                        TempSet(2) = TempMoves(M2)
                                        TempSet(3) = TempMoves(M3)
                                        TempSet(4) = TempMoves(M4)
                                        If Not Prevented(TempPKMN, TempSet) Then
                                            A = UBound(NowChecking) + 1
                                            ReDim Preserve NowChecking(A)
                                            ReDim Preserve CheckString(A)
                                            NowChecking(A) = X
                                            CheckString(A) = Temp
                                            B = Not BreedCheck(X, TempSet)
                                            Debug.Print "Check Finished, " & IIf(B, "Illegal", "Legal")
                                            ReDim Preserve NowChecking(0)
                                            ReDim Preserve CheckString(0)
'                                            If KamexCheck(.No, TempSet) = B Then
'                                                OtherArray = Split(Temp, "+")
'                                                For Z = 0 To UBound(OtherArray)
'                                                    OtherArray(Z) = Moves(OtherArray(Z)).Name
'                                                Next Z
'                                                Temp = Join(OtherArray, " / ")
'                                                Build = Build & IIf(B, "Illegal: ", "Legal: ") & vbNewLine & Temp
'                                            End If
                                            If B Then
                                                .BreedIllegals(Y) = .BreedIllegals(Y) & "|" & Temp
                                            End If
                                        End If
                                    Next M4
                                Next M3
                            Next M2
                        Next M1
                        
                        'Apply to all its evolutions
                        BasePKMN(X).BreedIllegals(Y) = .BreedIllegals(Y)
                        BasePKMN(X).DoneCheck = True
                        For M1 = 1 To 5
                            If .Evo(M1) > 0 Then
                                BasePKMN(.Evo(M1)).BreedIllegals(Y) = .BreedIllegals(Y)
                                BasePKMN(.Evo(M1)).DoneCheck = True
                            End If
                        Next M1
                    End If
                End If
            End With
            StepProgBar.Value = StepProgBar.Value + 1
            DoEvents
        Next X
    Next Y
    'Clipboard.Clear
    'Clipboard.SetText Build
    
'I'd use commenting instead, but VB keeps wanting to reset because of it...
GoTo SkipMe
    Build = ""
    A = 0
    For X = 1 To 386
        With BasePKMN(X)
            If .MyStage = 1 Then
                Build = Build & UCase(.Name) & vbNewLine & "--386 Illegals--"
                If .BreedIllegals(A) = "" Then
                    Build = Build & vbNewLine & "[None]"
                Else
                    TempArray = Split(.BreedIllegals(A), "|")
                    For Y = 1 To UBound(TempArray)
                        OtherArray = Split(TempArray(Y), "+")
                        For Z = 0 To UBound(OtherArray)
                            OtherArray(Z) = Moves(OtherArray(Z)).Name
                        Next Z
                        Temp = Join(OtherArray, " / ")
                        Build = Build & vbNewLine & Temp
                    Next Y
                End If
                Build = Build & vbNewLine & "--200 Illegals--"
                If .BreedIllegals(A + 1) = "" Then
                    Build = Build & vbNewLine & "[None]"
                Else
                    TempArray = Split(.BreedIllegals(A + 1), "|")
                    For Y = 1 To UBound(TempArray)
                        OtherArray = Split(TempArray(Y), "+")
                        For Z = 0 To UBound(OtherArray)
                            OtherArray(Z) = Moves(OtherArray(Z)).Name
                        Next Z
                        Temp = Join(OtherArray, " / ")
                        Build = Build & vbNewLine & Temp
                    Next Y
                End If
                Build = Build & vbNewLine & vbNewLine
            End If
        End With
        If X Mod 50 = 0 Then Beep
    Next X
    Clipboard.Clear
    Clipboard.SetText Build
SkipMe:

    lblStep.Caption = "Writing Pokémon Database"
    StepProgBar.Value = 0
    StepProgBar.Max = 389
    Open SlashPath & "PokeDB.csv" For Output As #1
    For X = 1 To 389
        With BasePKMN(X)
            Write #1, .No, .GSNo, .AdvNo, .Name, .Legendary, .Uber, .Type1, .Type2, .PAtt(0), .PAtt(1), .Color1, .Color2, .BaseHP, .BaseAttack, .BaseDefense, .BaseSpeed, .BaseSAttack, .BaseSDefense, .BaseSpecial, .StartsWith, .RawMoves, .RawMachine, .RawBreeding, .RawRBY, .RawRBYTM, .RawSpecial, .RawTutor, .RawAdv, .RawAdvTM, .RawAdvBreed, .RawAdvSpecial, .RawAdvTutor, .RawLFOnly, .ExistRBY, .ExistGSC, .ExistAdv, .PercentFemale, .RedBlue, .Yellow, .Gold, .Silver, .Crystal, .Ruby, .Sapphire, .MyStage, .MyMethod, .Evo(1), .EvoM(1), .Evo(2), .EvoM(2), .Evo(3), .EvoM(3), .Evo(4), .EvoM(4), .Evo(5), .EvoM(5), .Weight, .Height, .Offset, .LevelBal, .EggGroup1, .EggGroup2, .Illegals(0), .Illegals(1), .Illegals(2), .Illegals(3), .BreedIllegals(0), .BreedIllegals(1), .BreedIllegals(2), .BreedIllegals(3), .MoveLevels(0), .MoveLevels(1), .MoveLevels(2), .MoveLevels(3)
            StepProgBar.Value = StepProgBar.Value + 1
            DoEvents
        End With
    Next X
    Close
    
    lblStep.Caption = "Compressing Pokémon Database"
    DoEvents
    'PokeDB
    ReDim CBytes(FileLen(SlashPath & "PokeDB.csv") - 1) As Byte
    Open SlashPath & "PokeDB.csv" For Binary Access Read As #1
    Get #1, , CBytes()
    Close
    CompressZIt1.CompressData CBytes()
    If FileExists(SlashPath & "header.tmp") Then Kill SlashPath & "header.tmp"
    Open SlashPath & "header.tmp" For Output As #1
    Write #1, CompressZIt1.OriginalSize
    Close
    ReDim HBytes(FileLen(SlashPath & "header.tmp") - 1) As Byte
    Open SlashPath & "header.tmp" For Binary Access Read As #1
    Get #1, , HBytes()
    Close
    Kill SlashPath & "header.tmp"
    If FileExists(SlashPath & "PokeDB.cdb") Then Kill SlashPath & "PokeDB.cdb"
    Open SlashPath & "PokeDB.cdb" For Binary Access Write As #1
    Put #1, , HBytes
    Put #1, , CBytes
    Close
    
    T = Timer - T
    X = T \ 60
    T = Round(T - X * 60, 2)
    lblStep.Caption = "Finished in " & X & " minutes and " & T & " seconds."
    Command1.Visible = True
End Sub

Sub AddToArray(iArray() As String, ByVal Addition As String)
    ReDim Preserve iArray(UBound(iArray) + 1)
    iArray(UBound(iArray)) = Addition
End Sub
Function FileExists(ByVal FileName As String) As Boolean
    'Determines if a file exists
    On Error GoTo Failed
    If Dir(FileName) = "" Then FileExists = False Else FileExists = True
    Exit Function
Failed:
    FileExists = False
End Function

Private Function KamexCheck(PokeNum As Integer, MoveArray() As Integer) As Boolean
    'This checks for Breeding Move combinations that are illegal.
    'MAJOR thanks to Kamex for giving me the list (and major thanks
    'to my computer for processing that list into this code so I
    'wouldn't have to type all this myself =p)
    
    'If you're wondering about the freaky order, they're in order according
    'to their final evolution.  That's the way the list was given to me, meh.
    
    Select Case PokeNum
    Case 142 'AERODACTYL
        'Illegal: Foresight, Pursuit
        If HasMoves(MoveArray, 74, 154) Then KamexCheck = False: Exit Function
    Case 167, 168 'SPINARAK | ARIADOS
        'Illegal: Baton Pass, Sonicboom
        If HasMoves(MoveArray, 12, 198) Then KamexCheck = False: Exit Function
        'Illegal: Disable, Pursuit
        If HasMoves(MoveArray, 44, 154) Then KamexCheck = False: Exit Function
        'Illegal: Disable, Sonicboom
        If HasMoves(MoveArray, 44, 198) Then KamexCheck = False: Exit Function
        'Illegal: Psybeam, Pursuit
        If HasMoves(MoveArray, 150, 154) Then KamexCheck = False: Exit Function
        'Illegal: Psybeam, Sonicboom
        If HasMoves(MoveArray, 150, 198) Then KamexCheck = False: Exit Function
    Case 183, 184 'MARILL | AZUMARILL
        'Illegal: Amnesia, Foresight
        If HasMoves(MoveArray, 6, 74) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Present
        If HasMoves(MoveArray, 6, 148) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Supersonic
        If HasMoves(MoveArray, 6, 216) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Foresight
        If HasMoves(MoveArray, 14, 74) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Future Sight
        If HasMoves(MoveArray, 14, 79) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Present
        If HasMoves(MoveArray, 14, 148) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Supersonic
        If HasMoves(MoveArray, 14, 216) Then KamexCheck = False: Exit Function
        'Illegal: Foresight, Present
        If HasMoves(MoveArray, 74, 148) Then KamexCheck = False: Exit Function
        'Illegal: Foresight, Supersonic
        If HasMoves(MoveArray, 74, 216) Then KamexCheck = False: Exit Function
        'Illegal: Future Sight, Supersonic
        If HasMoves(MoveArray, 79, 216) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Perish Song
        If HasMoves(MoveArray, 109, 140) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Supersonic
        If HasMoves(MoveArray, 109, 216) Then KamexCheck = False: Exit Function
        'Illegal: Perish Song, Present
        If HasMoves(MoveArray, 140, 148) Then KamexCheck = False: Exit Function
        'Illegal: Perish Song, Supersonic
        If HasMoves(MoveArray, 140, 216) Then KamexCheck = False: Exit Function
        'Illegal: Present, Supersonic
        If HasMoves(MoveArray, 148, 216) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Future Sight, Light Screen
        If HasMoves(MoveArray, 6, 79, 109) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Future Sight, Perish Song
        If HasMoves(MoveArray, 6, 79, 140) Then KamexCheck = False: Exit Function
        'Illegal: Foresight, Future Sight, Perish Song
        If HasMoves(MoveArray, 74, 79, 140) Then KamexCheck = False: Exit Function
        'Illegal: Future Sight, Light Screen, Present
        If HasMoves(MoveArray, 79, 109, 148) Then KamexCheck = False: Exit Function
    Case 43, 44, 45, 182 'ODDISH | GLOOM | VILEPLUME | BELLOSSOM
        'Illegal: Flail, Swords Dance
        If HasMoves(MoveArray, 68, 222) Then KamexCheck = False: Exit Function
    Case 7, 8, 9 'SQUIRTLE | WARTORTLE | BLASTOISE
        'Illegal: Confusion, Flail
        If HasMoves(MoveArray, 29, 68) Then KamexCheck = False: Exit Function
        'Illegal: Confusion, Haze
        If HasMoves(MoveArray, 29, 87) Then KamexCheck = False: Exit Function
        'Illegal: Confusion, Mirror Coat
        If HasMoves(MoveArray, 29, 127) Then KamexCheck = False: Exit Function
        'Illegal: Confusion, Mist
        If HasMoves(MoveArray, 29, 129) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Foresight
        If HasMoves(MoveArray, 68, 74) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Haze
        If HasMoves(MoveArray, 68, 87) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Mirror Coat
        If HasMoves(MoveArray, 68, 127) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Mist
        If HasMoves(MoveArray, 68, 129) Then KamexCheck = False: Exit Function
        'Illegal: Foresight, Haze
        If HasMoves(MoveArray, 74, 87) Then KamexCheck = False: Exit Function
        'Illegal: Foresight, Mirror Coat
        If HasMoves(MoveArray, 74, 127) Then KamexCheck = False: Exit Function
        'Illegal: Haze, Mirror Coat
        If HasMoves(MoveArray, 87, 127) Then KamexCheck = False: Exit Function
    Case 113, 242 'CHANSEY | BLISSEY
        'Illegal: Heal Bell, Metronome
        If HasMoves(MoveArray, 89, 122) Then KamexCheck = False: Exit Function
    Case 4, 5, 6 'CHARMANDER | CHARMELEON | CHARIZARD
        'Illegal: Ancientpower, Beat Up
        If HasMoves(MoveArray, 7, 13) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Belly Drum
        If HasMoves(MoveArray, 7, 14) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Outrage
        If HasMoves(MoveArray, 7, 136) Then KamexCheck = False: Exit Function
        'Illegal: Beat Up, Belly Drum
        If HasMoves(MoveArray, 13, 14) Then KamexCheck = False: Exit Function
        'Illegal: Beat Up, Outrage
        If HasMoves(MoveArray, 13, 136) Then KamexCheck = False: Exit Function
        'Illegal: Beat Up, Rock Slide
        If HasMoves(MoveArray, 13, 167) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Bite
        If HasMoves(MoveArray, 14, 17) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Outrage
        If HasMoves(MoveArray, 14, 136) Then KamexCheck = False: Exit Function
    Case 35, 36, 173 'CLEFAIRY | CLEFABLE | CLEFFA
        'Illegal: Amnesia, Present
        If HasMoves(MoveArray, 6, 148) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Mimic
        If HasMoves(MoveArray, 14, 124) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Present
        If HasMoves(MoveArray, 14, 148) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Splash
        If HasMoves(MoveArray, 14, 204) Then KamexCheck = False: Exit Function
        'Illegal: Mimic, Present
        If HasMoves(MoveArray, 124, 148) Then KamexCheck = False: Exit Function
        'Illegal: Present, Splash
        If HasMoves(MoveArray, 148, 204) Then KamexCheck = False: Exit Function
    Case 90, 91 'SHELLDER | CLOYSTER
        'Illegal: Rapid Spin, Take Down
        If HasMoves(MoveArray, 158, 226) Then KamexCheck = False: Exit Function
    Case 222 'CORSOLA
        'Illegal: Amnesia, Rock Slide
        If HasMoves(MoveArray, 6, 167) Then KamexCheck = False: Exit Function
        'Illegal: Mist, Rock Slide
        If HasMoves(MoveArray, 129, 167) Then KamexCheck = False: Exit Function
        'Illegal: Mist, Screech
        If HasMoves(MoveArray, 129, 178) Then KamexCheck = False: Exit Function
        'Illegal: Rock Slide, Safeguard
        If HasMoves(MoveArray, 167, 173) Then KamexCheck = False: Exit Function
        'Illegal: Safeguard, Screech
        If HasMoves(MoveArray, 173, 178) Then KamexCheck = False: Exit Function
    Case 84, 85 'DODUO | DODRIO
        'Illegal: Faint Attack, Flail
        If HasMoves(MoveArray, 62, 68) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Haze, Quick Attack
        If HasMoves(MoveArray, 68, 87, 155) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Quick Attack, Supersonic
        If HasMoves(MoveArray, 68, 155, 216) Then KamexCheck = False: Exit Function
    Case 147, 148, 149 'DRATINI | DRAGONAIR | DRAGONITE
        'Illegal: Haze, Light Screen
        If HasMoves(MoveArray, 87, 109) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Mist
        If HasMoves(MoveArray, 109, 129) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Supersonic
        If HasMoves(MoveArray, 109, 216) Then KamexCheck = False: Exit Function
        'Illegal: Mist, Supersonic
        If HasMoves(MoveArray, 129, 216) Then KamexCheck = False: Exit Function
    Case 125, 239 'ELECTABUZZ | ELEKID
        'Illegal: Barrier, Cross Chop
        If HasMoves(MoveArray, 11, 36) Then KamexCheck = False: Exit Function
        'Illegal: Barrier, Karate Chop
        If HasMoves(MoveArray, 11, 103) Then KamexCheck = False: Exit Function
        'Illegal: Barrier, Rolling Kick
        If HasMoves(MoveArray, 11, 170) Then KamexCheck = False: Exit Function
        'Illegal: Barrier, Cross Chop, Meditate
        If HasMoves(MoveArray, 11, 36, 116) Then KamexCheck = False: Exit Function
    Case 102, 103 'EXEGGCUTE | EXEGGUTOR
        'Illegal: Ancientpower, Mega Drain
        If HasMoves(MoveArray, 7, 117) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Moonlight
        If HasMoves(MoveArray, 7, 130) Then KamexCheck = False: Exit Function
        'Illegal: Mega Drain, Moonlight, Synthesis
        If HasMoves(MoveArray, 117, 130, 223) Then KamexCheck = False: Exit Function
    Case 83 'FARFETCH'D
        'Illegal: Flail, Mirror Move
        If HasMoves(MoveArray, 68, 128) Then KamexCheck = False: Exit Function
    Case 21, 22 'SPEAROW | FEAROW
        'Illegal: Faint Attack, False Swipe
        If HasMoves(MoveArray, 62, 63) Then KamexCheck = False: Exit Function
        'Illegal: Faint Attack, Scary Face
        If HasMoves(MoveArray, 62, 176) Then KamexCheck = False: Exit Function
        'Illegal: False Swipe, Scary Face
        If HasMoves(MoveArray, 63, 176) Then KamexCheck = False: Exit Function
        'Illegal: False Swipe, Tri Attack
        If HasMoves(MoveArray, 63, 237) Then KamexCheck = False: Exit Function
        'Illegal: Quick Attack, Scary Face
        If HasMoves(MoveArray, 155, 176) Then KamexCheck = False: Exit Function
        'Illegal: Scary Face, Tri Attack
        If HasMoves(MoveArray, 176, 237) Then KamexCheck = False: Exit Function
    Case 158, 159, 160 'TOTODILE | CROCONAW | FERALIGATR
        'Illegal: Ancientpower, Crunch
        If HasMoves(MoveArray, 7, 37) Then KamexCheck = False: Exit Function
        'Illegal: Crunch, Razor Wind
        If HasMoves(MoveArray, 37, 160) Then KamexCheck = False: Exit Function
        'Illegal: Razor Wind, Rock Slide
        If HasMoves(MoveArray, 160, 167) Then KamexCheck = False: Exit Function
        'Illegal: Razor Wind, Thrash
        If HasMoves(MoveArray, 160, 229) Then KamexCheck = False: Exit Function
    Case 204, 205 'PINECO | FORRETRESS
        'Illegal: Flail, Pin Missile
        If HasMoves(MoveArray, 68, 142) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Reflect
        If HasMoves(MoveArray, 68, 162) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Swift
        If HasMoves(MoveArray, 68, 221) Then KamexCheck = False: Exit Function
    Case 92, 93, 94 'GASTLY | HAUNTER | GENGAR
        'Illegal: Haze, Perish Song
        If HasMoves(MoveArray, 87, 140) Then KamexCheck = False: Exit Function
    Case 207 'GLIGAR
        'Illegal: Counter, Razor Wind
        If HasMoves(MoveArray, 34, 160) Then KamexCheck = False: Exit Function
    Case 209, 210 'SNUBBULL | GRANBULL
        'Illegal: Crunch, Metronome
        If HasMoves(MoveArray, 37, 122) Then KamexCheck = False: Exit Function
        'Illegal: Faint Attack, Metronome
        If HasMoves(MoveArray, 62, 122) Then KamexCheck = False: Exit Function
        'Illegal: Heal Bell, Metronome
        If HasMoves(MoveArray, 89, 122) Then KamexCheck = False: Exit Function
        'Illegal: Leer, Metronome, Present
        If HasMoves(MoveArray, 107, 122, 148) Then KamexCheck = False: Exit Function
        'Illegal: Leer, Metronome, Reflect
        If HasMoves(MoveArray, 107, 122, 162) Then KamexCheck = False: Exit Function
        'Illegal: Metronome, Present, Reflect
        If HasMoves(MoveArray, 122, 148, 162) Then KamexCheck = False: Exit Function
    Case 214 'HERACROSS
        'Illegal: Bide, Flail, Harden
        If HasMoves(MoveArray, 15, 68, 86) Then KamexCheck = False: Exit Function
    Case 187, 188, 189 'HOPPIP | SKIPLOOM | JUMPLUFF
        'Illegal: Amnesia, Confusion
        If HasMoves(MoveArray, 6, 29) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Pay Day
        If HasMoves(MoveArray, 6, 138) Then KamexCheck = False: Exit Function
        'Illegal: Confusion, Encore
        If HasMoves(MoveArray, 29, 58) Then KamexCheck = False: Exit Function
        'Illegal: Confusion, Growl
        If HasMoves(MoveArray, 29, 82) Then KamexCheck = False: Exit Function
        'Illegal: Confusion, Pay Day
        If HasMoves(MoveArray, 29, 138) Then KamexCheck = False: Exit Function
        'Illegal: Encore, Pay Day
        If HasMoves(MoveArray, 58, 138) Then KamexCheck = False: Exit Function
    Case 140, 141 'KABUTO | KABUTOPS
        'Illegal: Aurora Beam, Dig
        If HasMoves(MoveArray, 9, 43) Then KamexCheck = False: Exit Function
        'Illegal: Bubblebeam, Flail
        If HasMoves(MoveArray, 24, 68) Then KamexCheck = False: Exit Function
        'Illegal: Aurora Beam, Flail, Rapid Spin
        If HasMoves(MoveArray, 9, 68, 158) Then KamexCheck = False: Exit Function
    Case 115 'KANGASKHAN
        'Illegal: Disable, Foresight
        If HasMoves(MoveArray, 44, 74) Then KamexCheck = False: Exit Function
        'Illegal: Focus Energy, Foresight
        If HasMoves(MoveArray, 73, 74) Then KamexCheck = False: Exit Function
        'Illegal: Focus Energy, Safeguard
        If HasMoves(MoveArray, 73, 173) Then KamexCheck = False: Exit Function
        'Illegal: Focus Energy, Stomp
        If HasMoves(MoveArray, 73, 207) Then KamexCheck = False: Exit Function
        'Illegal: Foresight, Stomp
        If HasMoves(MoveArray, 74, 207) Then KamexCheck = False: Exit Function
        'Illegal: Safeguard, Stomp
        If HasMoves(MoveArray, 173, 207) Then KamexCheck = False: Exit Function
    Case 116, 117, 230 'HORSEA | SEADRA | KINGDRA
        'Illegal: Disable, Dragon Rage
        If HasMoves(MoveArray, 44, 50) Then KamexCheck = False: Exit Function
        'Illegal: Disable, Flail
        If HasMoves(MoveArray, 44, 68) Then KamexCheck = False: Exit Function
        'Illegal: Disable, Octazooka
        If HasMoves(MoveArray, 44, 135) Then KamexCheck = False: Exit Function
        'Illegal: Disable, Splash
        If HasMoves(MoveArray, 44, 204) Then KamexCheck = False: Exit Function
        'Illegal: Dragon Rage, Octazooka
        If HasMoves(MoveArray, 50, 135) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Octazooka
        If HasMoves(MoveArray, 68, 135) Then KamexCheck = False: Exit Function
        'Illegal: Octazooka, Splash
        If HasMoves(MoveArray, 135, 204) Then KamexCheck = False: Exit Function
        'Illegal: Aurora Beam, Dragon Rage, Flail
        If HasMoves(MoveArray, 9, 50, 68) Then KamexCheck = False: Exit Function
        'Illegal: Aurora Beam, Dragon Rage, Splash
        If HasMoves(MoveArray, 9, 50, 204) Then KamexCheck = False: Exit Function
        'Illegal: Aurora Beam, Flail, Splash
        If HasMoves(MoveArray, 9, 68, 204) Then KamexCheck = False: Exit Function
    Case 98, 99 'KRABBY | KINGLER
        'Illegal: Amnesia, Dig
        If HasMoves(MoveArray, 6, 43) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Flail
        If HasMoves(MoveArray, 6, 68) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Haze
        If HasMoves(MoveArray, 6, 87) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Slam
        If HasMoves(MoveArray, 6, 187) Then KamexCheck = False: Exit Function
        'Illegal: Dig, Flail
        If HasMoves(MoveArray, 43, 68) Then KamexCheck = False: Exit Function
        'Illegal: Dig, Haze
        If HasMoves(MoveArray, 43, 87) Then KamexCheck = False: Exit Function
        'Illegal: Dig, Slam
        If HasMoves(MoveArray, 43, 187) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Haze
        If HasMoves(MoveArray, 68, 87) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Slam
        If HasMoves(MoveArray, 68, 187) Then KamexCheck = False: Exit Function
    Case 131 'LAPRAS
        'Illegal: Aurora Beam, Foresight
        If HasMoves(MoveArray, 9, 74) Then KamexCheck = False: Exit Function
    Case 108 'LICKITUNG
        'Illegal: Belly Drum, Magnitude
        If HasMoves(MoveArray, 14, 114) Then KamexCheck = False: Exit Function
        'Illegal: Body Slam, Magnitude
        If HasMoves(MoveArray, 19, 114) Then KamexCheck = False: Exit Function
    Case 66, 67, 68 'MACHOP | MACHOKE | MACHAMP
        'Illegal: Encore, Rolling Kick
        If HasMoves(MoveArray, 58, 170) Then KamexCheck = False: Exit Function
    Case 126, 240 'MAGMAR | MAGBY
        'Illegal: Barrier, Cross Chop
        If HasMoves(MoveArray, 11, 36) Then KamexCheck = False: Exit Function
        'Illegal: Barrier, Karate Chop
        If HasMoves(MoveArray, 11, 103) Then KamexCheck = False: Exit Function
        'Illegal: Cross Chop, Mega Punch, Screech
        If HasMoves(MoveArray, 36, 119, 178) Then KamexCheck = False: Exit Function
    Case 226 'MANTINE
        'Illegal: Haze, Hydro Pump, Slam, Twister
        If HasMoves(MoveArray, 87, 94, 187, 240) Then KamexCheck = False: Exit Function
    Case 104, 105 'CUBONE | MAROWAK
        'Illegal: Ancientpower, Belly Drum
        If HasMoves(MoveArray, 7, 14) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Perish Song
        If HasMoves(MoveArray, 7, 140) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Skull Bash
        If HasMoves(MoveArray, 7, 185) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Swords Dance
        If HasMoves(MoveArray, 7, 222) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Perish Song
        If HasMoves(MoveArray, 14, 140) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Swords Dance
        If HasMoves(MoveArray, 14, 222) Then KamexCheck = False: Exit Function
        'Illegal: Perish Song, Rock Slide
        If HasMoves(MoveArray, 140, 167) Then KamexCheck = False: Exit Function
        'Illegal: Perish Song, Screech
        If HasMoves(MoveArray, 140, 178) Then KamexCheck = False: Exit Function
        'Illegal: Perish Song, Swords Dance
        If HasMoves(MoveArray, 140, 222) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Rock Slide, Screech
        If HasMoves(MoveArray, 14, 167, 178) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Screech, Skull Bash
        If HasMoves(MoveArray, 14, 178, 185) Then KamexCheck = False: Exit Function
        'Illegal: Rock Slide, Screech, Swords Dance
        If HasMoves(MoveArray, 167, 178, 222) Then KamexCheck = False: Exit Function
    Case 152, 153, 154 'CHIKORITA | BAYLEEF | MEGANIUM
        'Illegal: Ancientpower, Counter
        If HasMoves(MoveArray, 7, 34) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Flail
        If HasMoves(MoveArray, 7, 68) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Swords Dance
        If HasMoves(MoveArray, 7, 222) Then KamexCheck = False: Exit Function
        'Illegal: Counter, Leech Seed
        If HasMoves(MoveArray, 34, 106) Then KamexCheck = False: Exit Function
        'Illegal: Counter, Vine Whip
        If HasMoves(MoveArray, 34, 242) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Leech Seed
        If HasMoves(MoveArray, 68, 106) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Swords Dance
        If HasMoves(MoveArray, 68, 222) Then KamexCheck = False: Exit Function
    Case 198 'MURKROW
        'Illegal: Drill Peck, Wing Attack
        If HasMoves(MoveArray, 53, 248) Then KamexCheck = False: Exit Function
    Case 163, 164 'HOOTHOOT | NOCTOWL
        'Illegal: Mirror Move, Supersonic
        If HasMoves(MoveArray, 128, 216) Then KamexCheck = False: Exit Function
        'Illegal: Faint Attack, Sky Attack, Supersonic
        If HasMoves(MoveArray, 62, 186, 216) Then KamexCheck = False: Exit Function
    Case 223, 224 'REMORAID | OCTILLERY
        'Illegal: Haze, Screech
        If HasMoves(MoveArray, 87, 178) Then KamexCheck = False: Exit Function
    Case 138, 139 'OMANYTE | OMASTAR
        'Illegal: Aurora Beam, Haze, Slam
        If HasMoves(MoveArray, 9, 87, 187) Then KamexCheck = False: Exit Function
        'Illegal: Aurora Beam, Slam, Supersonic
        If HasMoves(MoveArray, 9, 187, 216) Then KamexCheck = False: Exit Function
    Case 46, 47 'PARAS | PARASECT
        'Illegal: Counter, Psybeam
        If HasMoves(MoveArray, 34, 150) Then KamexCheck = False: Exit Function
        'Illegal: False Swipe, Flail
        If HasMoves(MoveArray, 63, 68) Then KamexCheck = False: Exit Function
        'Illegal: False Swipe, Psybeam
        If HasMoves(MoveArray, 63, 150) Then KamexCheck = False: Exit Function
        'Illegal: False Swipe, Screech
        If HasMoves(MoveArray, 63, 178) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Psybeam
        If HasMoves(MoveArray, 68, 150) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Pursuit
        If HasMoves(MoveArray, 68, 154) Then KamexCheck = False: Exit Function
        'Illegal: Flail, Screech
        If HasMoves(MoveArray, 68, 178) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Screech
        If HasMoves(MoveArray, 109, 178) Then KamexCheck = False: Exit Function
        'Illegal: Psybeam, Pursuit
        If HasMoves(MoveArray, 150, 154) Then KamexCheck = False: Exit Function
        'Illegal: Counter, Pursuit, Screech
        If HasMoves(MoveArray, 34, 154, 178) Then KamexCheck = False: Exit Function
    Case 16, 17, 18 'PIDGEY | PIDGEOTTO | PIDGEOT
        'Illegal: Foresight, Pursuit
        If HasMoves(MoveArray, 74, 154) Then KamexCheck = False: Exit Function
    Case 60, 61, 62, 186 'POLIWAG | POLIWHIRL | POLIWRATH | POLITOED
        'Illegal: Haze, Splash
        If HasMoves(MoveArray, 87, 204) Then KamexCheck = False: Exit Function
        'Illegal: Mist, Splash
        If HasMoves(MoveArray, 129, 204) Then KamexCheck = False: Exit Function
    Case 123, 212 'SCYTHER | SCIZOR
        'Illegal: Baton Pass, Counter
        If HasMoves(MoveArray, 12, 34) Then KamexCheck = False: Exit Function
        'Illegal: Baton Pass, Razor Wind
        If HasMoves(MoveArray, 12, 160) Then KamexCheck = False: Exit Function
        'Illegal: Baton Pass, Reversal
        If HasMoves(MoveArray, 12, 165) Then KamexCheck = False: Exit Function
        'Illegal: Counter, Razor Wind
        If HasMoves(MoveArray, 34, 160) Then KamexCheck = False: Exit Function
        'Illegal: Counter, Safeguard
        If HasMoves(MoveArray, 34, 173) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Razor Wind
        If HasMoves(MoveArray, 109, 160) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Reversal
        If HasMoves(MoveArray, 109, 165) Then KamexCheck = False: Exit Function
        'Illegal: Razor Wind, Reversal
        If HasMoves(MoveArray, 160, 165) Then KamexCheck = False: Exit Function
        'Illegal: Reversal, Safeguard
        If HasMoves(MoveArray, 165, 173) Then KamexCheck = False: Exit Function
    Case 118, 119 'GOLDEEN | SEAKING
        'Illegal: Hydro Pump, Psybeam
        If HasMoves(MoveArray, 94, 150) Then KamexCheck = False: Exit Function
    Case 79, 80, 199 'SLOWPOKE | SLOWBRO | SLOWKING
        'Illegal: Belly Drum, Future Sight
        If HasMoves(MoveArray, 14, 79) Then KamexCheck = False: Exit Function
        'Illegal: Belly Drum, Safeguard
        If HasMoves(MoveArray, 14, 173) Then KamexCheck = False: Exit Function
        'Illegal: Future Sight, Stomp
        If HasMoves(MoveArray, 79, 207) Then KamexCheck = False: Exit Function
        'Illegal: Safeguard, Stomp
        If HasMoves(MoveArray, 173, 207) Then KamexCheck = False: Exit Function
    Case 114 'TANGELA
        'Illegal: Amnesia, Confusion
        If HasMoves(MoveArray, 6, 29) Then KamexCheck = False: Exit Function
        'Illegal: Amnesia, Flail
        If HasMoves(MoveArray, 6, 68) Then KamexCheck = False: Exit Function
        'Illegal: Confusion, Flail
        If HasMoves(MoveArray, 29, 68) Then KamexCheck = False: Exit Function
    Case 72, 73 'TENTACOOL | TENTACRUEL
        'Illegal: Aurora Beam, Mirror Coat
        If HasMoves(MoveArray, 9, 127) Then KamexCheck = False: Exit Function
        'Illegal: Aurora Beam, Safeguard
        If HasMoves(MoveArray, 9, 173) Then KamexCheck = False: Exit Function
        'Illegal: Haze, Mirror Coat
        If HasMoves(MoveArray, 87, 127) Then KamexCheck = False: Exit Function
        'Illegal: Haze, Rapid Spin
        If HasMoves(MoveArray, 87, 158) Then KamexCheck = False: Exit Function
        'Illegal: Haze, Safeguard
        If HasMoves(MoveArray, 87, 173) Then KamexCheck = False: Exit Function
        'Illegal: Mirror Coat, Rapid Spin
        If HasMoves(MoveArray, 127, 158) Then KamexCheck = False: Exit Function
        'Illegal: Rapid Spin, Safeguard
        If HasMoves(MoveArray, 158, 173) Then KamexCheck = False: Exit Function
    Case 175, 176 'TOGEPI | TOGETIC
        'Illegal: Foresight, Present
        If HasMoves(MoveArray, 74, 148) Then KamexCheck = False: Exit Function
        'Illegal: Future Sight, Mirror Move
        If HasMoves(MoveArray, 79, 128) Then KamexCheck = False: Exit Function
        'Illegal: Mirror Move, Present
        If HasMoves(MoveArray, 128, 148) Then KamexCheck = False: Exit Function
        'Illegal: Peck, Present
        If HasMoves(MoveArray, 139, 148) Then KamexCheck = False: Exit Function
        'Illegal: Foresight, Future Sight, Peck
        If HasMoves(MoveArray, 74, 79, 139) Then KamexCheck = False: Exit Function
    Case 246, 247, 248 'LARVITAR | PUPITAR | TYRANITAR
        'Illegal: Ancientpower, Outrage
        If HasMoves(MoveArray, 7, 136) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Pursuit
        If HasMoves(MoveArray, 7, 154) Then KamexCheck = False: Exit Function
        'Illegal: Ancientpower, Stomp
        If HasMoves(MoveArray, 7, 207) Then KamexCheck = False: Exit Function
        'Illegal: Focus Energy, Outrage
        If HasMoves(MoveArray, 73, 136) Then KamexCheck = False: Exit Function
        'Illegal: Focus Energy, Pursuit
        If HasMoves(MoveArray, 73, 154) Then KamexCheck = False: Exit Function
        'Illegal: Focus Energy, Stomp
        If HasMoves(MoveArray, 73, 207) Then KamexCheck = False: Exit Function
        'Illegal: Outrage, Pursuit
        If HasMoves(MoveArray, 136, 154) Then KamexCheck = False: Exit Function
        'Illegal: Outrage, Stomp
        If HasMoves(MoveArray, 136, 207) Then KamexCheck = False: Exit Function
    Case 1, 2, 3 'BULBASAUR | IVYSAUR | VENUSAUR
        'Illegal: Light Screen, Razor Wind
        If HasMoves(MoveArray, 109, 160) Then KamexCheck = False: Exit Function
        'Illegal: Petal Dance, Razor Wind
        If HasMoves(MoveArray, 141, 160) Then KamexCheck = False: Exit Function
        'Illegal: Petal Dance, Skull Bash
        If HasMoves(MoveArray, 141, 185) Then KamexCheck = False: Exit Function
        'Illegal: Razor Wind, Safeguard
        If HasMoves(MoveArray, 160, 173) Then KamexCheck = False: Exit Function
        'Illegal: Razor Wind, Skull Bash
        If HasMoves(MoveArray, 160, 185) Then KamexCheck = False: Exit Function
        'Illegal: Light Screen, Safeguard, Skull Bash
        If HasMoves(MoveArray, 109, 173, 185) Then KamexCheck = False: Exit Function
    Case 69, 70, 71 'BELLSPROUT | WEEPINBELL | VICTREEBEL
        'Illegal: Encore, Leech Life
        If HasMoves(MoveArray, 58, 105) Then KamexCheck = False: Exit Function
        'Illegal: Encore, Swords Dance
        If HasMoves(MoveArray, 58, 222) Then KamexCheck = False: Exit Function
        'Illegal: Leech Life, Reflect, Synthesis
        If HasMoves(MoveArray, 105, 162, 223) Then KamexCheck = False: Exit Function
        'Illegal: Leech Life, Swords Dance, Synthesis
        If HasMoves(MoveArray, 105, 222, 223) Then KamexCheck = False: Exit Function
    Case 39, 40, 174 'JIGGLYPUFF | WIGGLYTUFF | IGGLYBUFF
        'Illegal: Faint Attack, Perish Song
        If HasMoves(MoveArray, 62, 140) Then KamexCheck = False: Exit Function
        'Illegal: Perish Song, Present
        If HasMoves(MoveArray, 140, 148) Then KamexCheck = False: Exit Function
    Case 193 'YANMA
        'Illegal: Leech Life, Reversal
        If HasMoves(MoveArray, 105, 165) Then KamexCheck = False: Exit Function
        'Illegal: Reversal, Whirlwind
        If HasMoves(MoveArray, 165, 247) Then KamexCheck = False: Exit Function
    End Select
    KamexCheck = True
End Function

