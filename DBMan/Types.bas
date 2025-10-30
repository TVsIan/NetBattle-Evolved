Attribute VB_Name = "Types"
Option Explicit
Option Compare Text

'Pokemon data
Public Type Pokemon
    No As Integer
    GSNo As Integer
    AdvNo As Integer
    Legendary As Boolean
    Uber As Boolean
    Image As String
    Name As String
    Nickname As String
    Type1 As Byte
    Type2 As Byte
    Attribute As Long
    PAtt(0 To 1) As Long
    AttNum As Byte
    NatureNum As Byte
    Color1 As Integer
    Color2 As Integer
    Move(1 To 4) As Integer
    MaxPP(1 To 4) As Byte
    PP(1 To 4) As Byte
    Item As Long
    Condition As Integer
    ConditionCount As Integer
    UnownLetter As Byte
    
    BaseHP As Integer
    BaseAttack As Integer
    BaseDefense As Integer
    BaseSpeed As Integer
    BaseSAttack As Integer
    BaseSDefense As Integer
    BaseSpecial As Integer
    
    MaxHP As Integer
    HP As Integer
    Attack As Integer
    Defense As Integer
    Speed As Integer
    SpecialAttack As Integer
    SpecialDefense As Integer
    
    DV_HP As Byte
    DV_Atk As Byte
    DV_Def As Byte
    DV_Spd As Byte
    DV_SAtk As Byte
    DV_SDef As Byte
    
    EV_HP As Byte
    EV_Atk As Byte
    EV_Def As Byte
    EV_Spd As Byte
    EV_SAtk As Byte
    EV_SDef As Byte
    
    Shiny As Boolean
    Level As Byte
    
    BaseMoves() As Integer
    MachineMoves() As Integer
    BreedingMoves() As Integer
    RBYMoves() As Integer
    RBYTM() As Integer
    SpecialMoves() As Integer
    MoveTutor() As Integer
    
    AdvMoves() As Integer
    ADVTM() As Integer
    AdvBreeding() As Integer
    AdvSpecial() As Integer
    AdvTutor() As Integer
    LFOnly() As Integer
    MoveLevels(0 To 3) As String
    ExistRBY As Boolean
    ExistGSC As Boolean
    ExistAdv As Boolean
    StartsWith As Byte
    PercentFemale As Integer
    Gender As Byte
    Evo(1 To 5) As Integer
    EvoM(1 To 5) As Integer
    Stage(1 To 5) As Integer
    MyStage As Integer
    MyMethod As Integer
    InBox As Integer
    GameVersion As CompatModes
    Weight As Integer
    Height As Integer
    Offset As Byte
    LevelBal As Byte
    RecycleItem As Long
    MarkerNum As Byte
    Illegals(3) As String
    BreedIllegals(3) As String
    EggGroup1 As Long
    EggGroup2 As Long
    
    'Rest in Slp/Frz Check
    Resting As Boolean

    'This is used for various berrys that involve randomness
    ItemEffect As Byte

    'This is the only really strange one - it sets to their position in your lineup.
    'Used for copying info from current to team.
    TeamNumber As Byte
    
    'Special Values for the Database
    RawMoves As String
    RawMachine As String
    RawBreeding As String
    RawRBY As String
    RawRBYTM As String
    RawSpecial As String
    RawTutor As String
    RawAdv As String
    RawAdvTM As String
    RawAdvBreed As String
    RawAdvSpecial As String
    RawAdvTutor As String
    RawLFOnly As String
    RedBlue As String
    Yellow As String
    Gold As String
    Silver As String
    Crystal As String
    Ruby As String
    Sapphire As String
    
    DoneCheck As Boolean
End Type
'Move stuff
Public Type Move
    'Need to use this one for IconList funkiness
    ID As Integer
    Name As String
    Type As Byte
    Power As Integer
    Accuracy As Byte
    PP As Byte
    'Text is a description - it comes up as a tooltip on the team builder
    'Note to self - add it as a tooltip on the battle screen
    Text As String
    SpecialPercent As Byte
    SpecialEffect As Byte
    WorksRight As Boolean
    BrightPowder As Boolean
    KingsRock As Boolean
    RBYMove As Boolean
    GSCMove As Boolean
    AdvMove As Boolean
    HitsTeam As Boolean
    SelfMove As Boolean
    OldTM As String
    NewTM As String
    ADVTM As String
    SubstituteBlocks As Boolean
    HitsAll As Boolean
    SoundMove As Boolean
    PhysMove As Boolean
    Target As Long
    MagicCoat As Boolean
End Type
Enum CompatModes
    nbRBYTrade
    nbGSCTrade
    nbTrueRuSa
    nbFullAdvance
    nbAdvTrades
    nbTrueRBY
    nbTrueGSC
End Enum

Enum MoveTypes
    nbRBYLevel = 1
    nbRBYTM
    nbGSCLevel
    nbGSCTM
    nbGSCEgg
    nbGSCTutor
    nbGSCSpecial
    nbAdvLevel
    nbAdvTM
    nbAdvEgg
    nbAdvTutor
    nbAdvSpecial
    nbAdvFL
End Enum


'Main database connection
Global PokeData As New ADODB.Connection

'BasePKMN = default values for Max Gene Pokemon
Global BasePKMN() As Pokemon

'Move information
Global Moves() As Move

'Type effectiveness chart - (AttackType,DefendType)
Global BattleMatrix(1 To 17, 1 To 17) As Single

Global SlashPath As String
Global Element(0 To 17) As String
Global AttributeText(77) As String
Global ColorText(10) As String
Global EvoMethod(0 To 15) As String
Global NowChecking() As Integer
Global CheckString() As String
Global CurrentPoke As Integer
Global Const InVBMode As Boolean = True
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)

Private Sub Main()
    If Right(App.Path, 1) = "\" Then SlashPath = App.Path Else SlashPath = App.Path & "\"
    PokeData.Provider = "Microsoft.Jet.OLEDB.4.0"
    PokeData.Properties("Data Source") = SlashPath & "PokeDB.mdb"
    PokeData.Properties("Jet OLEDB:Database Password") = "ginyu4ce"
    PokeData.Open
    'Descriptions for the elements
    Element(0) = ""
    Element(1) = "Normal"
    Element(2) = "Fire"
    Element(3) = "Water"
    Element(4) = "Electric"
    Element(5) = "Grass"
    Element(6) = "Ice"
    Element(7) = "Fighting"
    Element(8) = "Poison"
    Element(9) = "Ground"
    Element(10) = "Flying"
    Element(11) = "Psychic"
    Element(12) = "Bug"
    Element(13) = "Rock"
    Element(14) = "Ghost"
    Element(15) = "Dragon"
    Element(16) = "Dark"
    Element(17) = "Steel"
    'Color (Advance Pokedex) Text
    ColorText(0) = "(None)"
    ColorText(1) = "Green"
    ColorText(2) = "Red"
    ColorText(3) = "Blue"
    ColorText(4) = "Brown"
    ColorText(5) = "Yellow"
    ColorText(6) = "Purple"
    ColorText(7) = "Pink"
    ColorText(8) = "White"
    ColorText(9) = "Grey"
    ColorText(10) = "Black"
    'Attributes
    AttributeText(0) = "No Trait"
    AttributeText(1) = "Stench"
    AttributeText(2) = "Drizzle"
    AttributeText(3) = "Speed Boost"
    AttributeText(4) = "Battle Armor"
    AttributeText(5) = "Sturdy"
    AttributeText(6) = "Damp"
    AttributeText(7) = "Limber"
    AttributeText(8) = "Sand Veil"
    AttributeText(9) = "Static"
    AttributeText(10) = "Volt Absorb"
    AttributeText(11) = "Water Absorb"
    AttributeText(12) = "Oblivious"
    AttributeText(13) = "Cloud Nine"
    AttributeText(14) = "Compound Eyes"
    AttributeText(15) = "Insomnia"
    AttributeText(16) = "Color Change"
    AttributeText(17) = "Immunity"
    AttributeText(18) = "Flash Fire"
    AttributeText(19) = "Shield Dust"
    AttributeText(20) = "Own Tempo"
    AttributeText(21) = "Suction Cups"
    AttributeText(22) = "Intimidate"
    AttributeText(23) = "Shadow Tag"
    AttributeText(24) = "Rough Skin"
    AttributeText(25) = "Wonder Guard"
    AttributeText(26) = "Levitate"
    AttributeText(27) = "Effect Spore"
    AttributeText(28) = "Synchronize"
    AttributeText(29) = "Clear Body"
    AttributeText(30) = "Natural Cure"
    AttributeText(31) = "Lightning Rod"
    AttributeText(32) = "Serene Grace"
    AttributeText(33) = "Swift Swim"
    AttributeText(34) = "Chlorophyll"
    AttributeText(35) = "Illuminate"
    AttributeText(36) = "Trace"
    AttributeText(37) = "Huge Power"
    AttributeText(38) = "Poison Point"
    AttributeText(39) = "Inner Focus"
    AttributeText(40) = "Magma Armor"
    AttributeText(41) = "Water Veil"
    AttributeText(42) = "Magnet Pull"
    AttributeText(43) = "Soundproof"
    AttributeText(44) = "Rain Dish"
    AttributeText(45) = "Sand Stream"
    AttributeText(46) = "Pressure"
    AttributeText(47) = "Thick Fat"
    AttributeText(48) = "Early Bird"
    AttributeText(49) = "Flame Body"
    AttributeText(50) = "Run Away"
    AttributeText(51) = "Keen Eye"
    AttributeText(52) = "Hyper Cutter"
    AttributeText(53) = "Pickup"
    AttributeText(54) = "Truant"
    AttributeText(55) = "Hustle"
    AttributeText(56) = "Cute Charm"
    AttributeText(57) = "Plus"
    AttributeText(58) = "Minus"
    AttributeText(59) = "Forecast"
    AttributeText(60) = "Sticky Hold"
    AttributeText(61) = "Shed Skin"
    AttributeText(62) = "Guts"
    AttributeText(63) = "Marvel Scale"
    AttributeText(64) = "Liquid Ooze"
    AttributeText(65) = "Overgrow"
    AttributeText(66) = "Blaze"
    AttributeText(67) = "Torrent"
    AttributeText(68) = "Swarm"
    AttributeText(69) = "Rock Head"
    AttributeText(70) = "Drought"
    AttributeText(71) = "Arena Trap"
    AttributeText(72) = "Vital Spirit"
    AttributeText(73) = "White Smoke"
    AttributeText(74) = "Pure Power"
    AttributeText(75) = "Shell Armor"
    AttributeText(76) = "Cacophony"
    AttributeText(77) = "Air Lock"
    'Evolution
    EvoMethod(0) = "None"
    EvoMethod(1) = "Level"
    EvoMethod(2) = "Trade"
    EvoMethod(3) = "Thunder Stone"
    EvoMethod(4) = "Leaf Stone"
    EvoMethod(5) = "Water Stone"
    EvoMethod(6) = "Fire Stone"
    EvoMethod(7) = "Sun Stone"
    EvoMethod(8) = "Happiness"
    EvoMethod(9) = "Level (Day)"
    EvoMethod(10) = "Level (Night)"
    EvoMethod(11) = "Trade (With Item)"
    EvoMethod(12) = "Moon Stone"
    EvoMethod(13) = "Beauty"
    EvoMethod(14) = "Egg"
    EvoMethod(15) = "Egg (With Item)"
    Call LoadMoveData
    DBManMain.Show
    ReDim NowChecking(0)
    ReDim CheckString(0)
End Sub

Public Sub LoadPKMNData()
    'Load everything out of the database.
    'I'm not commenting everything because it should be fairly self-explanatory.
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    Dim TempVar As String
    Dim MTTemp As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim QueryResults As ADODB.Recordset
    Dim CurrentRecord As Integer
    
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Pokemon WHERE Number > 0 ORDER BY Number ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    While Not QueryResults.EOF
        CurrentRecord = QueryResults("Number")
        ReDim Preserve BasePKMN(CurrentRecord)
        If UBound(BasePKMN) < CurrentRecord Then
            ReDim Preserve BasePKMN(CurrentRecord) As Pokemon
        End If
        With BasePKMN(CurrentRecord)
            ReDim .BaseMoves(0)
            ReDim .MachineMoves(0)
            ReDim .BreedingMoves(0)
            ReDim .SpecialMoves(0)
            ReDim .RBYMoves(0)
            ReDim .MoveTutor(3)
            ReDim .RBYTM(0)
            ReDim .AdvMoves(0)
            ReDim .ADVTM(0)
            ReDim .AdvBreeding(0)
            ReDim .AdvSpecial(0)
            ReDim .AdvTutor(5)
            .No = CurrentRecord
            .Name = QueryResults("Name")
            .Type1 = QueryResults("Type1")
            If Len(QueryResults("Type2")) > 0 Then
                .Type2 = QueryResults("Type2")
            End If
            .BaseHP = QueryResults("HP")
            .BaseAttack = QueryResults("Attack")
            .BaseDefense = QueryResults("Defense")
            .BaseSpeed = QueryResults("Speed")
            .BaseSAttack = QueryResults("SpecialAttack")
            .BaseSDefense = QueryResults("SpecialDefense")
            .BaseSpecial = QueryResults("SpecialRBY")
            .StartsWith = QueryResults("BornWith")
            .PercentFemale = QueryResults("Percent Female")
            If Len(QueryResults("Moves")) > 0 Then
                Temp = QueryResults("Moves")
                Y = 1
                P1 = 0
                P2 = InStr(1, Temp, ",")
                While P2 > 0
                    ReDim Preserve .BaseMoves(UBound(.BaseMoves) + 1)
                    .BaseMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
                    P1 = P2
                    P2 = InStr(P1 + 1, Temp, ",")
                    Y = Y + 1
                Wend
            End If
            If Len(QueryResults("Machine Moves")) > 0 Then
                Temp = QueryResults("Machine Moves")
                Y = 1
                P1 = 0
                P2 = InStr(1, Temp, ",")
                While P2 > 0
                    ReDim Preserve .MachineMoves(UBound(.MachineMoves) + 1)
                    .MachineMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
                    P1 = P2
                    P2 = InStr(P1 + 1, Temp, ",")
                    Y = Y + 1
                Wend
            End If
            If Len(QueryResults("Breeding Moves")) > 0 Then
                Temp = QueryResults("Breeding Moves")
                Y = 1
                P1 = 0
                P2 = InStr(1, Temp, ",")
                While P2 > 0
                    ReDim Preserve .BreedingMoves(UBound(.BreedingMoves) + 1)
                    .BreedingMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
                    P1 = P2
                    P2 = InStr(P1 + 1, Temp, ",")
                    Y = Y + 1
                Wend
            End If
            If Len(QueryResults("R/B/Y Moves")) > 0 Then
                Temp = QueryResults("R/B/Y Moves")
                Y = 1
                P1 = 0
                P2 = InStr(1, Temp, ",")
                While P2 > 0
                    ReDim Preserve .RBYMoves(UBound(.RBYMoves) + 1)
                    .RBYMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
                    P1 = P2
                    P2 = InStr(P1 + 1, Temp, ",")
                    Y = Y + 1
                Wend
            End If
            If Len(QueryResults("Special Moves")) > 0 Then
                Temp = QueryResults("Special Moves")
                Y = 1
                P1 = 0
                P2 = InStr(1, Temp, ",")
                While P2 > 0
                    ReDim Preserve .SpecialMoves(UBound(.SpecialMoves) + 1)
                    .SpecialMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
                    P1 = P2
                    P2 = InStr(P1 + 1, Temp, ",")
                    Y = Y + 1
                Wend
            End If
            If Len(QueryResults("Move Tutor")) > 0 Then
                MTTemp = QueryResults("Move Tutor")
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
            'Fill in L100 stats for the Pokedex
            If Len(QueryResults("Evo1")) > 0 Then .Evo(1) = QueryResults("Evo1")
            If Len(QueryResults("Evo2")) > 0 Then .Evo(2) = QueryResults("Evo2")
            If Len(QueryResults("Evo3")) > 0 Then .Evo(3) = QueryResults("Evo3")
            If Len(QueryResults("Evo4")) > 0 Then .Evo(4) = QueryResults("Evo4")
            If Len(QueryResults("Evo5")) > 0 Then .Evo(5) = QueryResults("Evo5")
            If Len(QueryResults("EvoM1")) > 0 Then .EvoM(1) = QueryResults("EvoM1")
            If Len(QueryResults("EvoM2")) > 0 Then .EvoM(2) = QueryResults("EvoM2")
            If Len(QueryResults("EvoM3")) > 0 Then .EvoM(3) = QueryResults("EvoM3")
            If Len(QueryResults("EvoM4")) > 0 Then .EvoM(4) = QueryResults("EvoM4")
            If Len(QueryResults("EvoM5")) > 0 Then .EvoM(5) = QueryResults("EvoM5")
            If Len(QueryResults("Stage1")) > 0 Then .Stage(1) = QueryResults("Stage1")
            If Len(QueryResults("Stage2")) > 0 Then .Stage(2) = QueryResults("Stage2")
            If Len(QueryResults("Stage3")) > 0 Then .Stage(3) = QueryResults("Stage3")
            If Len(QueryResults("Stage4")) > 0 Then .Stage(4) = QueryResults("Stage4")
            If Len(QueryResults("Stage5")) > 0 Then .Stage(5) = QueryResults("Stage5")
            If Len(QueryResults("MyStage")) > 0 Then .MyStage = QueryResults("MyStage")
            If Len(QueryResults("MyMethod")) > 0 Then .MyMethod = QueryResults("MyMethod")
            .Legendary = QueryResults("Legendary")
            .Uber = QueryResults("Uber")
        End With
        Call ScanForDuplicates(CurrentRecord)
        QueryResults.MoveNext
    Wend
End Sub

Sub ScanForDuplicates(ByVal PKMN As Integer)
    'Clean up duplicate moves in the database
    Dim X As Integer
    Dim Y As Integer
    
    With BasePKMN(PKMN)
        'Move Tutor vs. All (except Level)
        For X = 1 To UBound(.MoveTutor)
            If .MoveTutor(X) > 0 Then
                For Y = 1 To UBound(.MachineMoves)
                    If .MachineMoves(Y) = .MoveTutor(X) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                Next
                For Y = 1 To UBound(.BreedingMoves)
                    If .BreedingMoves(Y) = .MoveTutor(X) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                Next
                For Y = 1 To UBound(.RBYMoves)
                    If .RBYMoves(Y) = .MoveTutor(X) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                Next
                For Y = 1 To UBound(.RBYTM)
                    If .RBYTM(Y) = .MoveTutor(X) Then .RBYTM(Y) = .RBYTM(Y) * -1
                Next
                For Y = 1 To UBound(.SpecialMoves)
                    If .SpecialMoves(Y) = .MoveTutor(X) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                Next
            End If
        Next
        'Base vs. All
        For X = 1 To UBound(.BaseMoves)
            If .BaseMoves(X) > 0 Then
                For Y = 1 To UBound(.MachineMoves)
                    If .MachineMoves(Y) = .BaseMoves(X) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                Next
                For Y = 1 To UBound(.BreedingMoves)
                    If .BreedingMoves(Y) = .BaseMoves(X) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                Next
                For Y = 1 To UBound(.RBYMoves)
                    If .RBYMoves(Y) = .BaseMoves(X) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                Next
                For Y = 1 To UBound(.RBYTM)
                    If .RBYTM(Y) = .BaseMoves(X) Then .RBYTM(Y) = .RBYTM(Y) * -1
                Next
                For Y = 1 To UBound(.SpecialMoves)
                    If .SpecialMoves(Y) = .BaseMoves(X) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                Next
                For Y = 1 To UBound(.MoveTutor)
                    If .MoveTutor(Y) = .BaseMoves(X) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                Next
            End If
        Next
        'Machine vs. All
        For X = 1 To UBound(.MachineMoves)
            If .MachineMoves(X) > 0 Then
                For Y = 1 To UBound(.BreedingMoves)
                    If .BreedingMoves(Y) = .MachineMoves(X) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                Next
                For Y = 1 To UBound(.RBYMoves)
                    If .RBYMoves(Y) = .MachineMoves(X) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                Next
                For Y = 1 To UBound(.RBYTM)
                    If .RBYTM(Y) = .MachineMoves(X) Then .RBYTM(Y) = .RBYTM(Y) * -1
                Next
                For Y = 1 To UBound(.SpecialMoves)
                    If .SpecialMoves(Y) = .MachineMoves(X) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                Next
            End If
        Next
        'Breeding/RBGY
        For X = 1 To UBound(.BreedingMoves)
            If .BreedingMoves(X) > 0 Then
                For Y = 1 To UBound(.RBYMoves)
                    If .RBYMoves(Y) = .BreedingMoves(X) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                Next
                For Y = 1 To UBound(.RBYTM)
                    If .RBYTM(Y) = .BreedingMoves(X) Then .RBYTM(Y) = .RBYTM(Y) * -1
                Next
                For Y = 1 To UBound(.SpecialMoves)
                    If .SpecialMoves(Y) = .BreedingMoves(X) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                Next
            End If
        Next
        'RBY/RBY Machines
        For X = 1 To UBound(.RBYMoves)
            If .RBYMoves(X) > 0 Then
                For Y = 1 To UBound(.RBYTM)
                    If .RBYTM(Y) = .RBYMoves(X) Then .RBYTM(Y) = .RBYTM(Y) * -1
                Next
                For Y = 1 To UBound(.SpecialMoves)
                    If .SpecialMoves(Y) = .RBYMoves(X) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                Next
            End If
        Next
    End With
End Sub

Public Sub LoadMoveData()
    'Load everything out of the database.
    'I'm not commenting everything because it should be fairly self-explanatory.
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    Dim TempVar As String
    Dim MTTemp As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim QueryResults As ADODB.Recordset
    Dim CurrentRecord As Integer
    
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Moves WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    While Not QueryResults.EOF
        CurrentRecord = QueryResults("ID")
        ReDim Preserve Moves(CurrentRecord)
        If UBound(Moves) < CurrentRecord Then
            ReDim Preserve Moves(CurrentRecord) As Move
        End If
        Moves(CurrentRecord).ID = CurrentRecord
        Moves(CurrentRecord).Name = QueryResults("Name")
        Moves(CurrentRecord).Type = QueryResults("Type")
        If Len(QueryResults("Power")) > 0 Then
            Moves(CurrentRecord).Power = QueryResults("Power")
        End If
        If Len(QueryResults("Accuracy")) > 0 Then
            Moves(CurrentRecord).Accuracy = QueryResults("Accuracy")
        End If
        Moves(CurrentRecord).PP = QueryResults("PP")
        If Len(QueryResults("Percent")) > 0 Then
            Moves(CurrentRecord).SpecialPercent = QueryResults("Percent")
        End If
        If Len(QueryResults("Special")) > 0 Then
            Moves(CurrentRecord).SpecialEffect = QueryResults("Special")
        End If
        Moves(CurrentRecord).Text = QueryResults("Description")
        Moves(CurrentRecord).WorksRight = QueryResults("Works Properly")
        Moves(CurrentRecord).BrightPowder = QueryResults("BrightPowder")
        Moves(CurrentRecord).KingsRock = QueryResults("KingsRock")
        Moves(CurrentRecord).RBYMove = QueryResults("RBYCompatible")
        Moves(CurrentRecord).GSCMove = QueryResults("GSCCompatible")
        Moves(CurrentRecord).AdvMove = QueryResults("ADVCompatible")
        Moves(CurrentRecord).SelfMove = QueryResults("AffectsSelf")
        Moves(CurrentRecord).Target = QueryResults("Target")
        If Len(QueryResults("RBYTM")) > 0 Then
            Moves(CurrentRecord).OldTM = QueryResults("RBYTM")
        End If
        If Len(QueryResults("GSTM")) > 0 Then
            Moves(CurrentRecord).NewTM = QueryResults("GSTM")
        End If
        If Len(QueryResults("AdvTM")) > 0 Then
            Moves(CurrentRecord).ADVTM = QueryResults("AdvTM")
        End If
        Moves(CurrentRecord).SubstituteBlocks = QueryResults("BlockSubstitute")
        QueryResults.MoveNext
        DoEvents
    Wend
End Sub
Public Sub SortStringArray(ByRef sArray() As String)
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   TriQuickSortString2 sArray, 4, iLBound, iUBound
   InsertionSortString sArray, iLBound, iUBound
End Sub
Private Sub TriQuickSortString2(ByRef sArray() As String, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2
      If sArray(iMin) > sArray(i) Then SwapStrings sArray(iMin), sArray(i)
      If sArray(iMin) > sArray(iMax) Then SwapStrings sArray(iMin), sArray(iMax)
      If sArray(i) > sArray(iMax) Then SwapStrings sArray(i), sArray(iMax)
      j = iMax - 1
      SwapStrings sArray(i), sArray(j)
      i = iMin
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4
      Do
         Do
            i = i + 1
         Loop While sArray(i) < sTemp
         Do
            j = j - 1
         Loop While sArray(j) > sTemp
         If j < i Then Exit Do
         SwapStrings sArray(i), sArray(j)
      Loop
      SwapStrings sArray(i), sArray(iMax - 1)
      TriQuickSortString2 sArray, iSplit, iMin, j
      TriQuickSortString2 sArray, iSplit, i + 1, iMax
   End If
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub
Private Sub InsertionSortString(ByRef sArray() As String, ByVal iMin As Long, ByVal iMax As Long)
   Dim i As Long
   Dim j As Long
   Dim sTemp As String
      For i = iMin + 1 To iMax
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4
      j = i
      Do While j > iMin
         If sArray(j - 1) <= sTemp Then Exit Do
         CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4
         j = j - 1
      Loop
      CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4
   Next i
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub
Private Sub SwapStrings(ByRef s1 As String, ByRef s2 As String)
   Dim i As Long
   i = StrPtr(s1)
   If i = 0 Then CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4
   CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
   CopyMemory ByVal VarPtr(s2), i, 4
End Sub

Public Function BreedCheck(PokeNum As Integer, MoveArray() As Integer) As Boolean
    Dim Mv() As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Temp As String
    Dim TempPKMN As Pokemon
    Dim Canceled As Boolean
    Dim ThisString As String
    Mv = MoveArray
    'BreedCheck = False: Exit Function
    
    With BasePKMN(PokeNum)
        'first off, the move is learned by anything other than
        'breeding, don't bother checking, any exceptions can be
        'entered in the database manually.
        For X = 1 To 4
            Y = Mv(X)
            If Y <> 0 Then
                If .GameVersion = nbGSCTrade Or .GameVersion = nbTrueGSC Then
                    If IsIn(.BaseMoves, Y) Or IsIn(.MachineMoves, Y) Or IsIn(.SpecialMoves, Y) Or IsIn(.MoveTutor, Y) Then Mv(X) = 0: Canceled = True
                    If .GameVersion = nbGSCTrade Then
                        If IsIn(.RBYMoves, Y) Or IsIn(.RBYTM, Y) Then Mv(X) = 0: Canceled = True
                    End If
                ElseIf .GameVersion = nbFullAdvance Or .GameVersion = nbTrueRuSa Then
                    If IsIn(.AdvMoves, Y) Or IsIn(.ADVTM, Y) Or IsIn(.AdvSpecial, Y) Then Mv(X) = 0: Canceled = True
                    If .GameVersion = nbFullAdvance Then
                        If IsIn(.AdvTutor, Y) Or IsIn(.LFOnly, Y) Then Mv(X) = 0: Canceled = True
                    End If
                End If
            End If
        Next X
        
        'Move all the blank moves to the end of the array
        For X = 1 To 4
            If Mv(X) = 0 Then
                Z = 0
                For Y = X To 3
                    If Mv(Y) <> 0 Then Z = 1
                    Mv(Y) = Mv(Y + 1)
                Next Y
                If Z = 1 Then X = X - 1 Else Exit For
            End If
        Next X
        
        'If we've already eliminated all the moves, exit here
        If X = 1 Then BreedCheck = True: Exit Function
        'debug.print UBound(NowChecking) & " Branches, Considering " & .Name & ": " & Moves(Mv(1)).Name & " " & Moves(Mv(2)).Name & " " & Moves(Mv(3)).Name & " " & Moves(Mv(4)).Name
        
        'If no moves were canceled and the BreedCheck is being called from
        'an Evolution, the check is pointless, exit to speed things up.
        ThisString = Mv(1)
        For X = 2 To 4
            If Mv(X) > 0 Then ThisString = ThisString & "+" & Mv(X)
        Next X
        If Not Canceled Then
            For Y = 1 To UBound(NowChecking) - 1
                For X = 1 To 5
                    If BasePKMN(NowChecking(Y)).Evo(X) = PokeNum Or NowChecking(Y) = PokeNum Then
                        If CheckString(Y) = ThisString Then
                            BreedCheck = False
                            Exit Function
                        End If
                    End If
                Next X
            Next Y
        End If
        CheckString(UBound(CheckString)) = ThisString
        Debug.Print UBound(NowChecking) & " Ply, Checking " & .Name & ": " & Moves(Mv(1)).Name & " " & Moves(Mv(2)).Name & " " & Moves(Mv(3)).Name & " " & Moves(Mv(4)).Name
        
        If .DoneCheck Then
            BreedCheck = Not Prevented(BasePKMN(PokeNum), Mv)
            Exit Function
        End If
    
        
        'Alright, now we go through each Poke and check:
        '1) Compatible with mode?
        '2) Compatible Egg Group?
        '3) Can it be a father?
        '4) Is the Poke already being checked?  (Avoids endless loops)
        '5) Can it get the moves in question legally?
        'If it passes all five, we've got a legal combo.
        X = IIf(.GameVersion = nbGSCTrade Or .GameVersion = nbTrueGSC, 251, 389)
        For X = 1 To X
            If Not (.GameVersion = nbTrueGSC And Not BasePKMN(X).ExistGSC) And Not (.GameVersion = nbTrueRuSa And Not BasePKMN(X).ExistAdv) Then
                If EggGroupCheck(.No, X) Then
                    If BasePKMN(X).PercentFemale < 15 Then
                        'If Not IsIn(NowChecking, X) Then
                            Y = UBound(NowChecking) + 1
                            ReDim Preserve NowChecking(Y)
                            ReDim Preserve CheckString(Y)
                            NowChecking(Y) = X
                            CheckString(Y) = ThisString
                            TempPKMN = BasePKMN(X)
                            For Z = 1 To 4
                                TempPKMN.Move(Z) = Mv(Z)
                            Next Z
                            Temp = LegalMove(TempPKMN, True)
                            ReDim Preserve NowChecking(UBound(NowChecking) - 1)
                            ReDim Preserve CheckString(UBound(CheckString) - 1)
                            If Temp = "" Then
                                Debug.Print "Legal Found: " & BasePKMN(X).Name
                                'CheckString(Y) = ""
                                BreedCheck = True
                                Exit Function
                            End If
                        'End If
                    End If
                End If
            End If
        Next X
        BreedCheck = False
        ReDim Preserve NowChecking(UBound(NowChecking) + 1)
        ReDim Preserve CheckString(UBound(CheckString) + 1)
        'debug.print BasePKMN(PokeNum).Name
    End With
End Function
Function EggGroupCheck(No1 As Integer, No2 As Integer, Optional SkipEvoCheck As Boolean = False) As Boolean
    Dim X1 As Integer
    Dim X2 As Integer
    Dim Y1 As Integer
    Dim Y2 As Integer
    If Not SkipEvoCheck Then
        On Error Resume Next
        For X1 = 0 To 5
            For X2 = 0 To 5
                Y1 = 0
                Y2 = 0
                If X1 = 0 Then
                    Y1 = No1
                Else
                    With BasePKMN(No1)
                        Y1 = IIf(BasePKMN(.Evo(X1)).MyStage > .MyStage, .Evo(X1), 0)
                    End With
                End If
                If X2 = 0 Then
                    Y2 = No2
                Else
                    With BasePKMN(No2)
                        Y2 = IIf(BasePKMN(.Evo(X2)).MyStage > .MyStage, .Evo(X2), 0)
                    End With
                End If
                If Y1 > 0 And Y2 > 0 Then
                    If EggGroupCheck(Y1, Y2, True) Then
                        EggGroupCheck = True
                        Exit Function
                    End If
                End If
            Next X2
        Next X1
    End If
    X1 = BasePKMN(No1).EggGroup1: X2 = BasePKMN(No1).EggGroup2
    Y1 = BasePKMN(No2).EggGroup1: Y2 = BasePKMN(No2).EggGroup2
    EggGroupCheck = False
    If X1 = Y1 And X1 > 0 And Y1 > 0 Then EggGroupCheck = True
    If X1 = Y2 And X1 > 0 And Y2 > 0 Then EggGroupCheck = True
    If X2 = Y1 And X2 > 0 And Y1 > 0 Then EggGroupCheck = True
    If X2 = Y2 And X2 > 0 And Y2 > 0 Then EggGroupCheck = True
End Function
Function IsIn(iArray() As Integer, Index As Integer) As Boolean
    Dim X As Long
    On Error GoTo Etrap
    For X = LBound(iArray) To UBound(iArray)
        If iArray(X) = Index Then IsIn = True: Exit Function
    Next X
Etrap:
    IsIn = False
End Function
Function Prevented(Poke As Pokemon, Check() As Integer)
    Dim TempArray() As String
    Dim TempCheck() As String
    Dim iString As String
    Dim iArray() As String
    Dim X As Integer
    Dim Y As Integer
    Select Case Poke.GameVersion
    Case nbTrueGSC
        iString = Poke.BreedIllegals(0) & Poke.BreedIllegals(1)
    Case nbGSCTrade
        iString = Poke.BreedIllegals(0)
    Case nbFullAdvance
        iString = Poke.BreedIllegals(2)
    Case nbTrueRuSa
        iString = Poke.BreedIllegals(2) & Poke.BreedIllegals(3)
    End Select
    iArray = Split(iString, "|")
    For X = 1 To UBound(iArray)
        TempArray = Split(iArray(X), "+")
        Select Case UBound(TempArray)
        Case 0: Prevented = HasMoves(Check, Val(TempArray(0)))
        Case 1: Prevented = HasMoves(Check, Val(TempArray(0)), Val(TempArray(1)))
        Case 2: Prevented = HasMoves(Check, Val(TempArray(0)), Val(TempArray(1)), Val(TempArray(2)))
        Case 3: Prevented = HasMoves(Check, Val(TempArray(0)), Val(TempArray(1)), Val(TempArray(2)), Val(TempArray(3)))
        End Select
        If Prevented Then Exit Function
    Next X
End Function
Function Factorial(n As Long) As Long
    If n = 0 Then Factorial = 1 Else Factorial = n * Factorial(n - 1)
End Function
