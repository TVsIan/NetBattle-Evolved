VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BattleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type OrderStructure
    MoveOrder As Byte
    EndTurnOrder As Byte
    MoveRandom As Byte
    EndTurnRandom As Byte
    PriorityTier As Integer
End Type

'Public properties
Public Player1 As Integer
Public Player2 As Integer
Public BattleString As String
Public DoMessages As Boolean
Public DoReplay As Boolean
Public Unrated As Boolean
Public RealWeather As WeatherTypes
Public TurnNumber As Integer
Public Terrain As Terrains
Public Waiting1 As Boolean
Public Waiting2 As Boolean
Public WatchIgnore1 As Boolean
Public WatchIgnore2 As Boolean
Public WatchFormID As Integer
Public WatchOK As Boolean
Public Timeout As Long
Public DisconStall As Boolean
Public TeamChecksum1 As String
Public TeamChecksum2 As String
Public WaitingFor As Byte
Public YourVersion As Byte 'For Watchers
Public StoredWinner As Byte
Public LogOutput As RTBClass
Public SendLastBStr As Boolean
Public SavedSyncStr As String

'Private info
Private NowWatching() As Integer
Private CheckedIn(1 To 2) As Boolean
Private Moved(1 To 4) As Boolean
Private Cancellable(1 To 4) As Boolean
Private SwitchTo(1 To 4) As Byte
Private RealRule(1 To 15) As Boolean
Private PName(1 To 2) As String
Private WeatherCount As Byte
Private GoesFirstOK As Boolean
Private GoesLastOK As Boolean
Private PKMN(1 To 2, 1 To 6) As Pokemon
Private Current(1 To 4) As Pokemon
Private BCondition(1 To 4) As BattleStuff
Private TeamBCondition(1 To 2) As TeamCond
Private BattlePauseNum As Byte
Private SwitchedThisTurn(1 To 4) As Boolean
Private StoreStadium(1 To 2, 1 To 3) As Byte
Private GotTeam(1 To 2) As Boolean
Private GameVersion(1 To 2) As Byte
Private Unfreezing(1 To 4) As Boolean
Private Order() As OrderStructure
Private BlankPos(1 To 2) As Byte
Public LogFile As String

'For rules and things
Public BattleMode As BattleModes 'RBY, GSC, or Adv
Public ActNum As Byte '2 if Single Battle, 4 if Double Battle

Public TextLog As String
Private DontTouchMe As Long

Public NumPoke As Byte
Public StadiumMode As Boolean
Public PresentRule As Boolean
Public RandBat As Boolean
Public NoWatch As Boolean
Public NoWatchChat As Boolean
Public UseTimeout As Boolean
Public PPUp As Boolean
Public ThisPNum As Byte
Public ModHash As String
Private MultiTarget As Boolean
Private SelfLoss(1 To 2) As Boolean
Private NeedSelect(1 To 2) As Boolean
Public BattleWindow As Battle

Public Function CastformImage(ByVal Pos As Byte) As Byte
    Select Case Current(Pos).Type1
    Case 1: CastformImage = 0
    Case 2: CastformImage = 2
    Case 3: CastformImage = 1
    Case 6: CastformImage = 4
    End Select
End Function

Public Function IsReady(Optional ByVal Team As Byte = 0) As Boolean
    If Team > 0 Then CheckedIn(Team) = True
    If CheckedIn(1) And CheckedIn(2) Then IsReady = True Else IsReady = False
End Function

Public Sub DoDamageMessage(ByVal Pos As Byte, ByVal Damage As Long)
    Dim X As Long
    Dim Y As Long
    X = Damage
    If TeamNum(Pos) <> ThisPNum And Not EnabledRule(nbExactHP) Then
        Y = Round((Current(Pos).HP + X) * 100 / Current(Pos).MaxHP)
        If Current(Pos).HP + X > 0 And Y = 0 Then Y = 1
        X = Round(CLng(Current(Pos).HP) * 100 / Current(Pos).MaxHP)
        If Current(Pos).HP > 0 And X = 0 Then X = 1
        X = Cap(Y - X, 100)
        Y = 1
    Else
        Y = 0
    End If
    If ActNum = 4 Then
        Call AddMessage(21 + Y, nbNumber, X, nbActive, Pos)
    Else
        Call AddMessage(17 + Y, nbNumber, X)
    End If
End Sub
Public Function NeedSwitch(Optional ByVal Team As Byte) As Boolean
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Dim B As Boolean
    NeedSwitch = False
    If BattleOver Then NeedSwitch = True: Exit Function
    For X = 1 To ActNum
        If Team = 0 Or Team = X Then
            If BCondition(X).BatonPassing Then NeedSwitch = True
            If Current(X).Condition = 8 And X <> BlankPos(TeamNum(X)) Then NeedSwitch = True
        End If
    Next X
End Function

Public Function FaintedPokes(Team As Byte) As Byte
    Dim X As Integer
    Dim Y As Byte
    Y = 0
    For X = 1 To NumPoke
        If PKMN(Team, X).Condition = 8 Or PKMN(Team, X).HP = 0 Then Y = Y + 1
    Next X
    FaintedPokes = Y
End Function


Public Function SetVer(ByVal Team As Byte, ByVal Value As Byte) As Boolean
    GameVersion(Team) = Value
    SetVer = True
End Function
Public Function GetVer(ByVal Team As Byte) As Byte
    GetVer = GameVersion(Team)
End Function

Public Property Get GotBothTeams() As Boolean
    If GotTeam(1) And GotTeam(2) Then GotBothTeams = True Else GotBothTeams = False
End Property

Public Property Get Rules() As String
    Rules = FixedHex(MakeBinArray(RealRule), 8)
End Property

Public Property Let PlayerName(ByVal Team As Long, ByVal NewName As String)
    PName(Team) = NewName
End Property
Public Property Get PlayerName(ByVal Team As Long) As String
    PlayerName = PName(Team)
End Property

'Return values:
'0 = Battle not over
'1 = Player 1
'2 = Player 2
'3 = Tie
Public Property Get Winner() As Byte
    Dim X As Byte
    Dim Fainted(2) As Byte
    If StoredWinner <> 0 Then
        Winner = StoredWinner
        Exit Property
    End If
    For X = 1 To NumPoke
        If PKMN(1, X).Condition = 8 Then Fainted(1) = Fainted(1) + 1
        If PKMN(2, X).Condition = 8 Then Fainted(2) = Fainted(2) + 1
    Next
    If Fainted(1) < NumPoke And Fainted(2) < NumPoke Then
        StoredWinner = 0
    ElseIf Fainted(1) < NumPoke And Fainted(2) = NumPoke Then
        StoredWinner = 1
    ElseIf Fainted(1) = NumPoke And Fainted(2) < NumPoke Then
        StoredWinner = 2
    ElseIf Fainted(1) = NumPoke And Fainted(2) = NumPoke Then
        If EnabledRule(nbSelfKO) And SelfLoss(2) Then
            StoredWinner = 1
        ElseIf EnabledRule(nbSelfKO) And SelfLoss(1) Then
            StoredWinner = 2
        Else
            StoredWinner = 3
        End If
    End If
    Winner = StoredWinner
End Property
Public Sub ClearSelfLoss()
    SelfLoss(1) = False
    SelfLoss(2) = False
End Sub
'Faints everybody if you manually set BattleOver
Public Property Let BattleOver(IsOver As Boolean)
    Dim X As Byte
    Dim Y As Byte
    
    If IsOver Then
        For X = 1 To 2
            Current(X).HP = 0
            Current(X).Condition = 8
            For Y = 1 To 6
                PKMN(X, Y).HP = 0
                PKMN(X, Y).Condition = 8
            Next
        Next
    End If
End Property

'Returns true if all Pokemon on at least one side are fainted
Public Property Get BattleOver() As Boolean
    Dim X As Byte
    Dim Fainted(2) As Byte
    
    For X = 1 To NumPoke
        If PKMN(1, X).Condition = 8 Then Fainted(1) = Fainted(1) + 1
        If PKMN(2, X).Condition = 8 Then Fainted(2) = Fainted(2) + 1
    Next
    If Fainted(1) = NumPoke Or Fainted(2) = NumPoke Then BattleOver = True Else BattleOver = False
End Property
Public Sub ForceLoss(ByVal Team As Byte)
    Dim Y As Integer
    Timeout = 0
    WaitingFor = 0
    Current(Team).HP = 0
    Current(Team).Condition = 8
    For Y = 1 To 6
        PKMN(Team, Y).HP = 0
        PKMN(Team, Y).Condition = 8
    Next
    DisconStall = False
End Sub
Public Function NeedsStadiumSelect(ByVal PNum As Byte) As Boolean
    NeedsStadiumSelect = NeedSelect(PNum)
End Function

'Player's Attack button should be enabled
Public Property Get CanAttack(ByVal PNum As Byte) As Boolean

    'Can't attack if current PKMN is fainted, turn is supposed to be skipped,
    'or battle is over.
    'Or if you're using Thrash, Outrage, or Petal Dance

    With BCondition(PNum)
        If NeedSwitch _
        Or .StuckMove <> 0 _
        Or (.BideCount <> 0 And BattleMode <> nbRBYBattle) _
        Or .Charging _
        Or .Recharging _
        Or .SemiInvul <> 0 _
        Or Me.BattleOver Then
            CanAttack = False
        Else
            CanAttack = True
        End If
    End With
End Property

'Player's Switch button should be enabled
Public Property Get CanSwitch(ByVal PNum As Byte) As Boolean
    Dim ONum As Byte
    
    If PNum = 1 Then ONum = 2 Else ONum = 1
    If Me.BattleOver Then CanSwitch = False: Exit Property
    'Absolutely enable switch if current PKMN is fainted, and it wasn't the last one.
    If NeedSwitch(PNum) Then CanSwitch = True: Exit Property
    'Disable if someone else needs a switch
    If NeedSwitch Then CanSwitch = False: Exit Property
    'Disable switch if the current PKMN is stuck in (Mean Look etc.), is doing a repeat
    'move (Thrash etc.), or the battle is over.
    With BCondition(PNum)
        If False _
        Or .StuckMove > 0 _
        Or (.BideCount <> 0 And Not BattleMode = nbRBYBattle) _
        Or .Recharging _
        Or .Charging _
        Or .SemiInvul <> 0 _
        Or .Ingrain _
        Or (.RepeatMove <> 0 And Not BattleMode = nbRBYBattle) _
        Or Trapped(PNum) _
        Or BCondition(ONum).BatonPassing Then
            CanSwitch = False
        Else
            CanSwitch = True
        End If
    End With
End Property
Public Function Trapped(Pos As Byte) As Boolean
    Dim X As Byte
    If BCondition(Pos).TrapNum > 0 Then Trapped = True
    For X = 1 To ActNum
        If X <> Pos And Current(X).HP > 0 Then
            If Current(X).Attribute = nbMagnetPull Then
                If TypeCheck(Current(Pos), nbSteel) Then Trapped = True
            End If
            If Current(X).Attribute = nbShadowTag And X <> AllyNum(Pos) Then Trapped = True
            If Current(X).Attribute = nbArenaTrap And X <> AllyNum(Pos) Then
                If Not TypeCheck(Current(Pos), nbFlying) And Current(Pos).Attribute <> nbLevitate Then Trapped = True
            End If
        End If
    Next
End Function
Public Function GetMoved(ByVal Pos As Long) As Boolean
    GetMoved = Moved(Pos)
End Function
Public Function GetSwitchTo(ByVal Pos As Long) As Byte
    GetSwitchTo = SwitchTo(Pos)
End Function

'Returns "" if the specified Poke can use the move in the specified slot.
'Returns the reason why if it can't.  Unless ReturnMessage is off, in which
'case it just returns "False"
Public Property Get CanUseMove(ByVal Pos As Byte, ByVal Slot As Byte, Optional ReturnMessage As Boolean = True) As String
    Dim X As Integer
    Dim Y As Byte
    Dim Z As Byte
    Dim Temp As String
    'Items that come first have the least priority.
    
    'Disable
    If Slot = BCondition(Pos).DisabledMove Then X = 7
    'Torment
    If BCondition(Pos).Torment And Slot = BCondition(Pos).LastMoveSlot Then X = 6
    'Taunt
    If BCondition(Pos).TauntCount > 0 And TauntForbids(Current(Pos).Move(Slot)) Then X = 5
    'Imprison
    For Y = OtherTeam(TeamNum(Pos)) To ActNum Step 2
        If BCondition(Y).Imprison Then
            For Z = 1 To 4
                If BCondition(Y).SealedMove(Z) = Current(Pos).Move(Slot) Then
                    X = 4
                    Exit For
                End If
            Next Z
        End If
    Next Y
    'Choice Band
    If Current(Pos).Item = nbChoiceBand And Slot <> BCondition(Pos).ChoiceBandSlot And BCondition(Pos).ChoiceBandSlot > 0 Then X = 3
    'Encore
    If BCondition(Pos).EncoreDuration > 0 And Slot <> BCondition(Pos).EncoreMove Then X = 2
    'PP Check
    If Current(Pos).PP(Slot) = 0 Then
        'RBY Wrap exception
        If BattleMode <> nbRBYBattle Or BCondition(Pos).RepeatMove = 0 Then X = 1
    End If

    
    If Not ReturnMessage Then
        If X > 0 Then CanUseMove = "False"
        Exit Property
    End If
    
    Select Case X
    Case 1
        CanUseMove = GetFText(311)
    Case 2
        Temp = Replace(GetFText(131), "%1", Moves(58).Name)
        CanUseMove = Replace(Temp, "%2", Moves(Current(Pos).Move(BCondition(Pos).EncoreMove)).Name)
    Case 3
        Temp = Replace(GetFText(131), "%1", Item(nbChoiceBand))
        CanUseMove = Replace(Temp, "%2", Moves(Current(Pos).Move(BCondition(Pos).ChoiceBandSlot)).Name)
    Case 4
        Temp = Replace(GetFText(53), "%1", Current(Pos).Nickname)
        CanUseMove = Replace(Temp, "%2", Moves(Current(Pos).Move(Slot)).Name)
    Case 5
        Temp = Replace(GetFText(52), "%1", Current(Pos).Nickname)
        CanUseMove = Replace(Temp, "%2", Moves(Current(Pos).Move(Slot)).Name)
    Case 6
        CanUseMove = Replace(GetFText(51), "%1", Current(Pos).Nickname)
    Case 7
        Temp = Replace(GetFText(50), "%1", Current(Pos).Nickname)
        CanUseMove = Replace(Temp, "%2", Moves(Current(Pos).Move(Slot)).Name)
    End Select
End Property
Public Function TauntForbids(M As Integer) As Boolean
    Dim X As Boolean
    X = (Moves(M).Power > 0)
    Select Case M
    Case 15, 34, 50, 67, 68, 75, 84, 91, 93, 112, 114, 127, 133, 148, 153, 164, 165, 179, 198, 215, 275, 327
        X = True
    End Select
    TauntForbids = Not X
End Function

Public Property Get StruggleOK(ByVal Pos As Byte) As Boolean
    Dim X As Byte
    Dim Y As Byte
    For X = 1 To 4
        If Current(Pos).Move(X) > 0 Then
            If Me.CanUseMove(Pos, X, False) = "" Then Y = Y + 1
        End If
    Next X
    StruggleOK = (Y = 0)
End Property


Public Function SetSPoke(ByVal Team As Byte, ByVal Data As String) As Boolean
    On Error GoTo Failed
    StoreStadium(Team, 1) = Val(ChopString(Data, 1))
    StoreStadium(Team, 2) = Val(ChopString(Data, 1))
    StoreStadium(Team, 3) = Val(Data)
    NeedSelect(Team) = False
    SetSPoke = True
    Exit Function
Failed:
    SetSPoke = False
End Function

'Generates a string containing a Pokemon team's data
Public Function GetTeam(ByVal Team As Byte) As String
    Dim Temp As String
    Dim Temp2 As String
    Dim X As Byte
    For X = 1 To 6
        Temp = Temp & PKMN2Str(PKMN(Team, X))
    Next
    'Current(Team) = PKMN(Team, 1)
    'Debug.Print "GetTeam Len: " & Len(Temp)
    GetTeam = Temp
End Function

'Interprets a GetTeam-generated string & initializes Pokemon
Public Function SetTeam(ByVal Team As Byte, _
    ByVal TString As String) As Boolean
    Dim X As Byte
    On Error GoTo Failed
    If Len(TString) <> 210 Then
        SetTeam = False
        If DebugMode Then MsgBox "Invalid data received!", vbCritical, "Error"
        Exit Function
    End If

    For X = 1 To 6
        PKMN(Team, X) = Str2PKMN(ChopString(TString, POKELEN))
        With PKMN(Team, X)
            Select Case BattleMode
            Case nbRBYBattle: .GameVersion = nbRBYTrade
            Case nbGSCBattle: .GameVersion = nbGSCTrade
            Case nbAdvBattle: If .GameVersion <> nbModAdv Then .GameVersion = nbFullAdvance
            End Select
            If EnabledRule(nbLevelBalance) Then .Level = .LevelBal
            Call FillInPokeData(PKMN(Team, X), .GameVersion)
            Call SetPokePP(Team, X, 1, PPUp)
            Call SetPokePP(Team, X, 2, PPUp)
            Call SetPokePP(Team, X, 3, PPUp)
            Call SetPokePP(Team, X, 4, PPUp)
            .Image = ChooseImage(PKMN(Team, X), GameVersion(Team))
            .Condition = 1
            .TeamNumber = X
            If Not DoMessages And Not DoReplay Then .Nickname = ServerDB.WordFilter(.Nickname)
        End With
    Next
    GotTeam(Team) = True
    SetTeam = True
    Exit Function
Failed:
    SetTeam = False
End Function

'Load a Pokemon into a team position from Pokemon format data (for loading own team)
Public Function SetPKMN(ByVal Team As Byte, ByVal Position As Byte) As Boolean
    Dim X As Byte
    
    On Error GoTo Failed
    PKMN(Team, Position) = Code.SwapClassPKMN
    With PKMN(Team, Position)
        If EnabledRule(nbLevelBalance) Then .Level = .LevelBal
        If .Nickname = "" Then .Nickname = .Name
        .Condition = 1
        .TeamNumber = Position
        For X = 1 To 4
            Call SetPokePP(Team, Position, X, PPUp)
        Next
        If Position = 1 Then Current(Team) = PKMN(Team, Position)
        .MaxHP = GetHP(.Level, .BaseHP, .DV_HP)
        .HP = .MaxHP
        .Attack = GetStat(.Level, .BaseAttack, .DV_Atk)
        .Defense = GetStat(.Level, .BaseDefense, .DV_Def)
        .Speed = GetStat(.Level, .BaseSpeed, .DV_Spd)
        If BattleMode = nbRBYBattle Then
            .Item = nbNoItem
            .SpecialAttack = GetStat(.Level, .BaseSpecial, .DV_SAtk)
            .SpecialDefense = GetStat(.Level, .BaseSpecial, .DV_SAtk)
        Else
            .SpecialAttack = GetStat(.Level, .BaseSAttack, .DV_SAtk)
            .SpecialDefense = GetStat(.Level, .BaseSDefense, .DV_SAtk)
        End If
    End With
    SetPKMN = True
    Exit Function
Failed:
    SetPKMN = False
End Function

Public Property Let Rules(ByVal Hexified As String)
    Dim RuleData As Integer
    Dim X As Byte
    Dim Y As Byte
        
    RuleData = Dec(Hexified)
    Call Code.ReadBinArray(RuleData, RealRule())
    'If Player1 > 0 Then BattleMode = CompatVersion(Player(Player1).GameVersion)
    If EnabledRule(nbUsePPUps) Then PPUp = True Else PPUp = False
    'ActNum is set via the server now
    'ActNum = IIf(EnabledRule(nbDoubleBattle), 4, 2)
    If EnabledRule(nbStadiumMode) Then
        NumPoke = 3
        StadiumMode = True
        NeedSelect(1) = True
        NeedSelect(2) = True
    Else
        NumPoke = 6
        NeedSelect(1) = False
        NeedSelect(2) = False
    End If
    If EnabledRule(nbRandbat) Then RandBat = True Else RandBat = False
    If EnabledRule(nbNoWatch) Then NoWatch = True Else NoWatch = False
'    If EnabledRule(nbNoWatchChat) Then NoWatchChat = True Else NoWatchChat = False
    If EnabledRule(nbTimeout) Then UseTimeout = True Else UseTimeout = False
    If EnabledRule(nbUnrated) Then Unrated = True Else Unrated = False
    If EnabledRule(nbPresentRule) Then PresentRule = True Else PresentRule = False
    If Not BattleWindow Is Nothing Then
        With BattleWindow
            If EnabledRule(nbUnrated) Then .mnuBattleItem(2).Enabled = False
            If EnabledRule(nbExactHP) Then
                .HPBar(.ONum).Caption = nbExact
                If ActNum = 4 Then .HPBar(.ONum + 2).Caption = nbExact
            End If
        End With
    End If

    ReDim Order(1 To ActNum)
    If Not BattleWindow Is Nothing Then
        If BattleWindow.Resuming Then Exit Property
    End If
    
    Call AddMessage(312 + BattleMode)
    For X = 1 To UBound(RealRule)
        If RealRule(X) Then Call AddMessage(23, nbRule, X)
    Next
End Property
Public Function EnabledRule(RuleNum As RuleEnum) As Boolean
    EnabledRule = RealRule(RuleNum)
End Function

'Get a PKMN's data in Pokemon format
'0 = current, 1-6 = Team PKMN
Public Function SetSwapPKMN(ByVal Team As Byte, ByVal Position As Byte) As String
    Select Case Position
        Case 0
            If Current(Team).No = 0 Then Current(Team) = PKMN(Team, 1)
            Code.SwapClassPKMN = Current(Team)
        Case 1 To 6
            Code.SwapClassPKMN = PKMN(Team, Position)
    End Select
End Function
Public Function SetSwapBC(ByVal Team As Byte) As String
    Code.SwapBCondition = BCondition(Team)
End Function
Public Function SetSwapTC(ByVal Team As Byte) As String
    Code.SwapBTeamCond = TeamBCondition(Team)
End Function
Public Sub SetItemEffect(Effects() As Byte)
    Dim X As Integer
    For X = 1 To 6
        PKMN(1, X).ItemEffect = Effects(X)
    Next X
    For X = 1 To 6
        PKMN(2, X).ItemEffect = Effects(X + 6)
    Next X
End Sub
Public Sub SetTrace(Number As Integer)
    Dim Build As String
    Dim X As Byte
    Build = Dec2Bin(Number, 4)
    For X = 1 To ActNum
        BCondition(X).TraceNumber = Val(Mid(Build, X, 1))
    Next X
End Sub
'Parse BattleString and do the operations
Public Function ParseBattle(ByVal NewBString As String) As Boolean
    Static BString As String
    Dim X As Byte
    Dim Y As Integer
    Dim Command(1 To 2) As String
    Dim Data(1 To 2) As String
    Dim Temp As String
    Dim AtkPos As Byte
    Dim DefPos As Byte
    Dim M As Integer
    Dim Hit As Boolean
    Dim HitCount As Byte
    Dim Damage() As Long
    Dim CHit As Byte
    Dim EffectTook As Integer
    Dim EffectParam As Integer
    Dim CHitArray() As Boolean
    Dim GoAhead As Boolean
    Dim ConditionEffect As Integer
    Dim CalledBy As Integer
    Dim XMess As Byte
    Dim SwitchPoke As Byte
    Dim SkipMess As Byte
    Dim DamageString As String
    Dim FullString As String
    Dim FocusBand As Boolean
    Dim TraitEffect() As Byte
    
    On Error GoTo Failed
    BattleWindow.tmrDelay.Enabled = True
    If BString = "" Then
        'battlewindow.tmrDelay.Enabled = True
        'battlewindow.tmrBattleQueue.Enabled = False
        If NewBString = "" Then BattleWindow.tmrDelay.Enabled = False: Exit Function
        
        BString = Chr2Bin(NewBString) & "0"
    
        For X = 1 To ActNum
            Temp = ChopString(BString, 1)
            If TeamNum(X) <> BattleWindow.PNum Or BattleWindow.ImJustWatching Then
                SwitchTo(X) = Val(Temp)
            End If
        Next X
    
        If Not NeedSwitch Then
            If BMessStyle = 1 Then Call AddMessage(32, nbNumber, TurnNumber)
            For X = 1 To ActNum
                With BCondition(X)
                    .Protected = False
                    .Enduring = False
                    .MagicCoat = False
                    .InPursuit = False
                    .HelpingHand = False
                    .Snatch = 0
                    .Soundproofed = 0
                End With
                Order(X).MoveRandom = Bin2Dec(ChopString(BString, 2)) + 1
                Order(X).EndTurnRandom = Bin2Dec(ChopString(BString, 2)) + 1
                Order(X).PriorityTier = Bin2Dec(ChopString(BString, 4)) - 5
            Next X
            For X = 1 To ActNum
                If ChopString(BString, 1) = "1" Then
                    Call AddMessage(263, nbActive, X)
                End If
            Next X
            For X = 1 To ActNum
                BCondition(X).ShedSkin = (ChopString(BString, 1) = "1")
            Next X
        End If
    End If
    
    '>>> Call WriteDebugLog("Battle data: " & BString)
    While Len(BString) <> 0
        Temp = ChopString(BString, 1)
        Select Case Temp
            Case "1" 'Attack
                Erase Damage
                Erase CHitArray
                Erase TraitEffect
                For Y = 1 To ActNum
                    BCondition(Y).Targeted = CBool(ChopString(BString, 1))
                Next Y
                AtkPos = Bin2Dec(ChopString(BString, 2)) + 1
                M = Bin2Dec(ChopString(BString, 9))
                CalledBy = Bin2Dec(ChopString(BString, 3))
                EffectTook = Bin2Dec(ChopString(BString, 10))
                ConditionEffect = Bin2Dec(ChopString(BString, 5))
                For Y = 10 To 2 Step -1
                    BCondition(AtkPos).MoveUsed(Y) = BCondition(AtkPos).MoveUsed(Y - 1)
                Next Y
                BCondition(AtkPos).MoveUsed(1) = M
                Call DoCountdown(AtkPos)
                Call DoPreTurn(AtkPos, M, CalledBy, EffectTook, ConditionEffect, GoAhead)
                If GoAhead Then
                    Do
                        For Y = 1 To ActNum
                            If BCondition(Y).Targeted Then
                                Hit = CBool(ChopString(BString, 1))
                                HitCount = Bin2Dec(ChopString(BString, 3))
                                EffectTook = Bin2Dec(ChopString(BString, 4))
                                EffectParam = Bin2Dec(ChopString(BString, 10))
                                FocusBand = CBool(ChopString(BString, 1))
                                For X = 1 To HitCount
                                    ReDim Preserve Damage(1 To X)
                                    ReDim Preserve CHitArray(1 To X)
                                    ReDim Preserve TraitEffect(1 To X)
                                    Damage(X) = Bin2Dec(ChopString(BString, 10))
                                    CHitArray(X) = Bin2Bool(ChopString(BString, 1))
                                    TraitEffect(X) = Bin2Dec(ChopString(BString, 4))
                                Next X
                                Call DoTurn(AtkPos, Y, M, Hit, EffectTook, EffectParam, FocusBand, Damage, CHitArray, TraitEffect)
                                BCondition(Y).Targeted = False
                                Exit For
                            End If
                        Next Y
                    Loop Until Y = ActNum + 1
                End If
                Call SyncPKMN
                If Not NeedSwitch And BattleMode <> nbAdvBattle And Not BCondition(AtkPos).InPursuit Then Call DoConditionsGSC(AtkPos)
                If BMessStyle = 1 And Not NeedSwitch And BattlePauseNum <= 1 And (Not BCondition(AtkPos).InPursuit Or BattleMode = nbGSCBattle) Then Call AddMessage(315)
                If MoveDelay > 0 And Not BattleWindow.SkipDelay And Not NeedSwitch Then
                    Call BattleWindow.UpdateImages
                    Call BattleWindow.UpdateStats
                    Exit Function
'                    Call battlewindow.HPAnimTimer_Timer
'                    DoEvents
'                    SleepIt MoveDelay
                End If
            Case "0" 'Switch
                'The BString will have extra 0's attached, that's what all the checks are for.
                If Len(BString) >= 6 Then
                    AtkPos = Bin2Dec(ChopString(BString, 2)) + 1
                    SwitchPoke = Bin2Dec(ChopString(BString, 3))
                    If SwitchPoke > 0 Then
                        BCondition(AtkPos).TraceNumber = Bin2Dec(ChopString(BString, 1))
                        Call DoSwitch(AtkPos, SwitchPoke, BCondition(AtkPos).BatonPassing)
                        BCondition(AtkPos).BatonPassing = False
                        If BattleMode <> nbAdvBattle Then
                            If BCondition(OtherTeam(AtkPos)).InPursuit Then Call DoConditionsGSC(OtherTeam(AtkPos))
                        End If
                        If BMessStyle = 1 And Not NeedSwitch And BattlePauseNum <= 1 Then Call AddMessage(315)
                        
                        
                        If MoveDelay > 0 And Not BattleWindow.SkipDelay And Not NeedSwitch Then
                            Call BattleWindow.UpdateImages
                            Call BattleWindow.UpdateStats
                            Exit Function
'                            DoEvents
'                            SleepIt MoveDelay
                        End If
                    End If
                Else
                    BString = ""
                End If
        End Select
    Wend
    
    BattleWindow.tmrBattleQueue.Enabled = True
    BattleWindow.tmrDelay.Enabled = False
    Call BattleWindow.AddToQueue("ENDTN")
    
    If BattlePauseNum = 0 Then BattlePauseNum = 1
    If Not NeedSwitch Then
        If BattleMode = nbAdvBattle Then
            Call DoEndRoundAdv(BattlePauseNum)
        Else
            Call DoEndRoundGSC(BattlePauseNum)
        End If
        If BattlePauseNum = 7 Then BattlePauseNum = 1
    End If

    '>>> Call WriteDebugLog("String parsed properly.")
    ParseBattle = True
    'SleepIt 5000
    Exit Function
Failed:
    MsgBox "Run Time Error " & Err.Number & vbNewLine & Err.Description, vbCritical, "Error"
    '>>> Call WriteDebugLog("Error parsing battle string: remaining text is " & BString)
    If InVBMode Then
        On Error GoTo 0
        Resume
    End If
    ParseBattle = False
End Function
Private Function SleepIt(Milliseconds As Long)
    Dim X As Long
    For X = 1 To Milliseconds * 1000
        'Sleep 1
        DoEvents
    Next X
End Function

'Add a move to the queue
Public Function LoadMove(ByVal Team As Byte, ByVal MoveNum As Integer, TargNum As Byte, Optional SkipQueue As Boolean = False) As Boolean
    
    On Error GoTo Failed
    Dim X As Byte
    Dim Fainted As Byte
    Dim RealMove As Integer
    
    If Me.BattleOver Then LoadMove = True: Exit Function
    If Not Me.CanAttack(Team) Then LoadMove = False: Exit Function
    
    'Shift previous moves back
    If (BattleMode <> nbRBYBattle Or (Current(Team).Condition <> nbSlp And Current(Team).Condition <> nbFrz)) And Not SkipQueue Then
        For X = 1 To 9
            BCondition(Team).MoveUsed(X + 1) = BCondition(Team).MoveUsed(X)
        Next
    
        'Make MoveUsed(1) the selected move
        '0 = Struggle
        '1-4 = PKMN's moves
        'Anything else - L33T H4X0RING!  Return False and bug out.
        'Also, check the PP of the selected move.
        'If zero - ALSO L33T H4X0RING!
        Select Case MoveNum
            Case 0
                BCondition(Team).MoveUsed(1) = 210
            Case 1 To 4
                RealMove = Current(Team).Move(MoveNum)
                If RealMove = 0 And InVBMode Then Stop
                If CanUseMove(Team, MoveNum, False) <> "" Then LoadMove = False: Exit Function
                With BCondition(Team)
                    If .EncoreMove <> 0 And .EncoreMove <> MoveNum Then LoadMove = False: Exit Function
                    If .StuckMove <> 0 And .StuckMove <> RealMove Then LoadMove = False: Exit Function
                    .MoveUsed(1) = RealMove
                End With
            Case Else
                LoadMove = False
                Exit Function
        End Select
        If Not BCondition(Team).Encore Then BCondition(Team).TargetNum = TargNum
    End If
    'Move processed.
    Moved(Team) = True
    
    'All set!  Make sure to check .Ready from ServerWindow
    LoadMove = True
    Exit Function
    
Failed:
    'Return False if failed for some reason.
    LoadMove = False
End Function

'Add a Pokemon switch to the queue
Public Function LoadSwitch(ByVal Pos As Byte, ByVal PokeNum As Integer) As Boolean
    Dim X As Integer
    Dim Team As Byte
    Dim Blank As Pokemon
    On Error GoTo Failed
    Team = TeamNum(Pos)
    If Me.BattleOver Then LoadSwitch = True: Exit Function

    'Fail if it was a bad switch
    If PKMN(Team, PokeNum).HP <= 0 _
    Or PokeNum = Current(Pos).TeamNumber _
    Or PokeNum = Current(AllyNum(Pos)).TeamNumber _
    Or Pos = BlankPos(Team) _
    Or Not Me.CanSwitch(Pos) Then
        LoadSwitch = False
        Exit Function
    Else
        For X = Team To ActNum Step 2
            If SwitchTo(X) = PokeNum Then
                LoadSwitch = False
                Exit Function
            End If
        Next X
        SwitchTo(Pos) = PokeNum
    End If
    'All set!  Make sure to check .Ready from ServerWindow
    LoadSwitch = True
    Exit Function
Failed:
    'Return False if failed for some reason.
    LoadSwitch = False
End Function

'User cancelled the move
Public Sub UnloadMove(ByVal Pos As Long)
    Moved(Pos) = False
    Call WriteDebugLog("Move Unloaded")
End Sub
'User cancelled the switch
Public Sub UnloadSwitch(ByVal Pos As Long)
    SwitchTo(Pos) = 0
End Sub
Public Sub StruggleDump(ByVal Pos As Long)

End Sub
'Check to see if the battle is ready to proceed
'Scans a single player if PNum is used.
Public Property Get Ready(Optional ByVal PNum As Long = 0) As Boolean
    Dim X As Byte
    Dim Y As Byte
    Dim NS(1 To 4) As Boolean
    Dim B As Boolean
    For X = 1 To ActNum
        NS(X) = NeedSwitch(X)
        If NS(X) Then B = True
    Next X
    
    For X = 1 To ActNum
        If PNum = 0 Or PNum = X Then
        
            If Not CheckedIn(TeamNum(X)) Then Ready = False: Exit Property
            
            If NS(X) And SwitchTo(X) = 0 Then
                'see footnote----
                If ActNum = 4 Then
                    Y = AllyNum(X)
                    If NS(Y) And SwitchTo(Y) <> 0 Then
                        Y = OtherTeam(TeamNum(X))
                        If NS(Y) Or NS(Y + 2) Then
                            Ready = False: Exit Property
                        End If
                    Else
                        Ready = False: Exit Property
                    End If
                Else
                    Ready = False: Exit Property
                End If
                '----------------
            ElseIf Not B And Current(X).HP > 0 Then
                If Not Moved(X) _
                And SwitchTo(X) = 0 _
                And Not BCondition(X).Charging _
                And Not BCondition(X).Recharging _
                And BCondition(X).SemiInvul = 0 _
                And (BCondition(X).BideCount = 0 Or BattleMode = nbRBYBattle) _
                And BCondition(X).StuckMove = 0 _
                Then Ready = False: Exit Property
            End If
        End If
    Next X
    Ready = True
    
    'Ok, the reason for this crazyness is a strange attribute of
    'Ruby/Sapp's fainting mechanics.  Apperently, if more than one
    'Pokémon is fainted at any given time, here's how it works: If
    'pokes are fainted on both teams, then the game waits for both
    'players to choose replacement pokes.  However, two allied pokes
    'are down and the opposing two are both fine, then the switches
    'happen one by one as soon as the player selected each poke.  So,
    'yeah, that's what this is doing there.  Meh...
End Property


'The following functions and subroutines set the order for that each pokemon
'will move in.  It works with both Single Battles and Double Battles.
'
'This looks kinda complicated for something this simple, let me explain
'why.  Obviously, there are now potentially 4 speeds to order instead of
'just 2, but what makes it even worse is that the order in which End Round
'things such as Perish Song countdown and Sandstorm/Hail damage are executed
'now (dun dun DUN!) in order of Pokemon speed!  Not only that, you know how
'two equal speeds are just ordered randomly?  Well, the random numbers are
'different for the move order and the End Round order, meaning basically things
'have to be calculated twice.  In addition to THAT, since there is randomness
'involved, all this has to be sent to the clients as well.  Ugh.
Private Function GetPokemonOrder() As Byte()
    'This one gets the basic order by each Pokemon's Speed
    Dim X As Byte
    Dim Y As Integer
    Dim Speed() As Integer
    Dim SpeedBuild() As Integer
    Dim MainBuild() As Byte
    Dim Temp() As Byte
    ReDim Speed(1 To ActNum)
    ReDim SpeedBuild(1 To ActNum)
    ReDim MainBuild(1 To ActNum)
    For X = 1 To ActNum
        Speed(X) = GetSpd(X)
    Next
    For X = 1 To ActNum
        For Y = 1 To ActNum - 1
            If Speed(X) > SpeedBuild(Y) Then Exit For
        Next Y
        For Y = ActNum To Y + 1 Step -1
            SpeedBuild(Y) = SpeedBuild(Y - 1)
            MainBuild(Y) = MainBuild(Y - 1)
        Next Y
        SpeedBuild(Y) = Speed(X)
        MainBuild(Y) = X
    Next X
    GetPokemonOrder = MainBuild
End Function
Function GetRandomOrder() As Byte()
    'This gets a random order for use in randomizing equal Speeds.
    Dim W As Byte
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Dim Build() As Byte
    Dim Taken() As Boolean
    ReDim Build(1 To ActNum)
    ReDim Taken(1 To ActNum)
    For X = 1 To ActNum
        Z = Int(Rnd * (ActNum - X + 1)) + 1
        For Y = 1 To ActNum
            If Not Taken(Y) Then Z = Z - 1
            If Z = 0 Then Exit For
        Next Y
        Build(X) = Y
        Taken(Y) = True
    Next X
    GetRandomOrder = Build
End Function
Sub SetPriorityTiers(MoveEffect() As Integer)
    Dim X As Byte
    If BattleMode = nbAdvBattle Then
        For X = 1 To ActNum
            Select Case MoveEffect(X)
            Case 144 'Helping Hand
                Order(X).PriorityTier = 5
            Case 148, 165 'Magic Coat, Snatch
                Order(X).PriorityTier = 4
            Case 31, 41, 141 'Protect, Detect, Endure, Follow Me
                Order(X).PriorityTier = 3
            Case 43, 137 'Quick Attack, Extremespeed, Mach Punch, Fake Out
                Order(X).PriorityTier = 2
            '--------------------------------------.
            Case 120 'Vital Throw                  |
                Order(X).PriorityTier = -1 '       |
            Case 140 'Focus Punch                  |
                Order(X).PriorityTier = -2 '       |
            Case 159 'Revenge                      |
                Order(X).PriorityTier = -3 '       |
            Case 26, 72 'Counter, Mirror Coat      |
                Order(X).PriorityTier = -4 '       |
            Case 93 'Roar, Whirlwind               |
                Order(X).PriorityTier = -5 '       |
                                           '       |
                                           '       |
            Case Else 'Quick Claw >----------------`
                If Current(X).Item = nbQuickClaw Then
                    Order(X).PriorityTier = IIf(Odds(23.4), 1, 0)
                Else  'Everything Else
                    Order(X).PriorityTier = 0
                End If
            End Select
        Next X
    Else
        For X = 1 To 2
            Select Case MoveEffect(X)
            Case 31, 41 'Protect, Detect, Endure
                Order(X).PriorityTier = 3
            Case 43 'Quick Attack, Extremespeed, Mach Punch
                Order(X).PriorityTier = 2
            '-----------------------------------------.
            Case 26 'Counter                          |
                Order(X).PriorityTier = -1 '          |
            Case 72, 93, 120 'Vital Throw, Roar,      |
                             'Whirlwind, Mirror Coat  |
                             '(non-RBY only)          |
                If BattleMode <> nbRBYBattle Then '   |
                    Order(X).PriorityTier = -1 '      |
                End If                     '          |
                                           '          |
                                           '          |
            Case Else 'Quick Claw >-------------------`
                If Current(X).Item = nbQuickClaw Then
                    Order(X).PriorityTier = IIf(Odds(23.4), 1, 0)
                Else  'Everything Else
                    Order(X).PriorityTier = 0
                End If
            End Select
        Next X
    End If
End Sub
Function GetOrderArray(Optional ForMoves As Boolean = False) As Byte()
    'This puts it all together into an array.
    Dim X As Byte
    Dim Y As Integer
    Dim Z As Byte
    Dim A As Integer
    Dim Randoms() As Byte
    Dim BaseOrder() As Byte
    Dim Build() As Byte
    Dim TempBuild() As Byte
    Dim TempRandom() As Byte
    Dim RandOrder() As Byte
    Dim Speed() As Integer
    ReDim Randoms(1 To ActNum)
    ReDim BaseOrder(1 To ActNum)
    ReDim Build(1 To ActNum)
    ReDim Speed(1 To ActNum)
    For X = 1 To ActNum
        Speed(X) = GetSpd(X)
        If ForMoves Then
            Randoms(X) = Order(X).MoveRandom
        Else
            Randoms(X) = Order(X).EndTurnRandom
        End If
    Next X
    BaseOrder = GetPokemonOrder
    
    'Apply the randoms
    For X = 1 To ActNum - 1
        For Y = X + 1 To ActNum
            If Speed(BaseOrder(X)) <> Speed(BaseOrder(Y)) Then Exit For
        Next Y
        Y = Y - 1
        If Y <> X Then
            ReDim TempBuild(X To Y)
            ReDim RandOrder(X To Y)
            For Z = X To Y
                For A = X To Y - 1
                    If Randoms(Z) > TempBuild(A) Then Exit For
                Next A
                For A = Y To A + 1 Step -1
                    TempBuild(A) = TempBuild(A - 1)
                    RandOrder(A) = RandOrder(A - 1)
                Next A
                TempBuild(A) = Randoms(Z)
                RandOrder(A) = BaseOrder(Z)
            Next Z
            For Z = X To Y
                BaseOrder(Z) = RandOrder(Z)
            Next Z
        End If
    Next X
    
    'Apply the priority tiers
    ReDim TempBuild(1 To ActNum)
    A = 0
    For Y = 5 To -5 Step -1
        For X = 1 To ActNum
            If Order(BaseOrder(X)).PriorityTier = Y Then
                A = A + 1
                TempBuild(A) = BaseOrder(X)
            End If
        Next X
    Next Y
    GetOrderArray = TempBuild
End Function

'Call this when ready
Public Sub DoBattle()
    Dim MoveOrder() As Byte
    Dim MoveEffect() As Integer
    Dim Resuming As Boolean
    Dim PursuitOK As Boolean
    Dim ProtectOK As Boolean
    Dim X As Byte
    Dim Y As Integer
    Dim Z As Byte
    
    ''Debug.Print vbNewLine & "BATTLE OUTPUT"
    ReDim MoveOrder(1 To ActNum)
    ReDim MoveEffect(1 To ActNum)
    BattleString = ""
    
    For X = 1 To ActNum
        If SwitchTo(X) > 0 Then
            Moved(X) = False
            BattleString = BattleString & "1"
        Else
            BattleString = BattleString & "0"
        End If
    Next X
    
    If ActNum = 2 Then
        BCondition(1).TargetNum = 2
        BCondition(2).TargetNum = 1
    End If
    
    If BattlePauseNum = 0 Then
        
        'Reset everything
        If Not NeedSwitch Then
            For X = 1 To ActNum
                If Current(X).TeamNumber = 0 Then Current(X).TeamNumber = 1
                SwitchedThisTurn(X) = False
                Resuming = False
                GoesFirstOK = False
                GoesLastOK = False
                With BCondition(X)
                    If BattleMode <> nbRBYBattle Then
                        .LastPhDamage = 0
                        .LastSpDamage = 0
                        .LastPhHitter = 0
                        .LastSpHitter = 0
                    End If
                    .InPursuit = False
                    .BeingPursued = False
                    .Protected = False
                    .Enduring = False
                    .Flinching = False
                    .Damaged = False
                    .HelpingHand = False
                    .FollowMe = False
                    .MagicCoat = False
                    .Snatch = 0
                    .Soundproofed = 0
                    If .SemiInvul <> 0 Then Moved(X) = True
                    If .Recharging Then Moved(X) = True
                    If .Charging Then Moved(X) = True
                    If .BideCount > 0 Then Moved(X) = True
                    If .StuckMove > 0 Then Moved(X) = True
                    If BattleMode = nbRBYBattle Then
                        If BCondition(OtherTeam(X)).RepeatMove > 0 Then Moved(X) = False
                    End If
                    MoveEffect(X) = Moves(.MoveUsed(1)).SpecialEffect
                End With
            Next X
            MoveOrder = GetRandomOrder
            For X = 1 To ActNum
                Order(X).MoveRandom = MoveOrder(X)
            Next X
            MoveOrder = GetRandomOrder
            For X = 1 To ActNum
                Order(X).EndTurnRandom = MoveOrder(X)
                MoveOrder(X) = 0
            Next X
        End If
        
        For X = 1 To ActNum
            If SwitchTo(X) > 0 Then MoveEffect(X) = 0
        Next
        
        'Figure the speed
        Call SetPriorityTiers(MoveEffect)
        MoveOrder = GetOrderArray(True)

        'Communicate the order to the client
        For X = 1 To ActNum
            BattleString = BattleString & Dec2Bin(Order(X).MoveRandom - 1, 2)
            BattleString = BattleString & Dec2Bin(Order(X).EndTurnRandom - 1, 2)
            BattleString = BattleString & Dec2Bin(Order(X).PriorityTier + 5, 4)
        Next X
        
        'Focus Punch is weird
        For X = 1 To ActNum
            If MoveEffect(X) = 140 And Current(X).Condition <> 4 And Current(X).Condition <> 7 And Moved(X) Then
                BattleString = BattleString & "1"
            Else
                BattleString = BattleString & "0"
            End If
        Next X
        
        'For Shed Skin
        For X = 1 To ActNum '____________________________
            Y = IIf(Odds(30, Int(Rnd * 256)), "1", "0") '|
            BCondition(X).ShedSkin = CBool(Y) '          |
            BattleString = BattleString & CStr(Y) '      |
        Next X '                                         |
    'The reason for this is that I don't want the the True Random
    'system pulling 4 numbers out of the queue every single turn
    'when most of the time they won't even be used.  The thing is,
    'the ShedSkin variable has to be set even if the Poke doesn't
    'have Shed Skin.  The reason is that the Poke may obtain Shed
    'Skin via Skill Swap or Roll Play before the activation check
    'in DoConditionsAdv, so the program has to know whether or not
    'it should activate if this situation comes up.
    Else
        MoveOrder = GetOrderArray(True)
    End If
    
    '***SWITCHES***
    For X = 1 To ActNum
        If SwitchTo(X) > 0 Then
            If Not BCondition(X).BatonPassing Then
                Z = OtherTeam(TeamNum(X))
                If ActNum = 4 Then Z = Z + 2
                For Y = Z To 1 Step -2
                    PursuitOK = (MoveEffect(Y) = 86 And BCondition(Y).TargetNum = X And Moved(Y))
                    If BattleMode = nbAdvBattle Then
                        If Current(Y).Condition = 4 Or Current(Y).Condition = 7 Or (Current(Y).Attribute = nbTruant And BCondition(Y).Truant) Then PursuitOK = False
                    End If
                    If PursuitOK Then
                        BCondition(Y).InPursuit = True
                        Call DoPreAttack(Y)
                        If NeedSwitch(X) Then Exit For
                    End If
                Next Y
            End If
            BCondition(X).TraceNumber = GetSmallRnd(1)
            Call DoSwitch(X, SwitchTo(X), BCondition(X).BatonPassing)
            Moved(X) = False
            BattleString = BattleString & "0" & Dec2Bin(X - 1, 2) & Dec2Bin(Current(X).TeamNumber, 3) & CStr(BCondition(X).TraceNumber)
            BCondition(X).BatonPassing = False
            If BattleMode <> nbAdvBattle Then
                If BCondition(OtherTeam(X)).InPursuit Then Call DoConditionsGSC(OtherTeam(X))
            End If
            If NeedSwitch(X) Then Exit For
        End If
    Next X
    
    '***ATTACKS***
    For Y = 1 To ActNum
        X = MoveOrder(Y)
        If BattleMode = nbAdvBattle Then
            GoesFirstOK = False
            For Z = 1 To ActNum
                If Z <> X And Not SwitchedThisTurn(Z) Then GoesFirstOK = True
            Next Z
            GoesLastOK = True
        Else
            GoesFirstOK = (Y = 1 And Not BCondition(OtherTeam(X)).SwitchedThisTurn)
            GoesLastOK = (Y = 2)
        End If
        If Moved(X) And Not NeedSwitch Then
            Call DoCountdown(X)
            Call DoPreAttack(X)
            If NeedSwitch Then
                Exit For
            Else
                If BattleMode <> nbAdvBattle Then Call DoConditionsGSC(X)
            End If
        End If
    Next Y
    
    If BattlePauseNum = 0 Then BattlePauseNum = 1
    If Not NeedSwitch Then
        If BattleMode = nbAdvBattle Then
            Call DoEndRoundAdv(BattlePauseNum)
        Else
            Call DoEndRoundGSC(BattlePauseNum)
        End If
        If BattlePauseNum = 7 Then BattlePauseNum = 0
    End If

    BattleString = "BCMD:" & Bin2Chr(BattleString)
    
    SendLastBStr = DisconStall
    
    For X = 1 To 2
        Y = OtherTeam(X)
        CheckedIn(X) = ((NeedSwitch(Y) Or NeedSwitch(Y + 2)) And Not (NeedSwitch(X) Or NeedSwitch(X + 2)))
    Next X
End Sub

Sub ReOrder(OriginalOrder() As Byte, MoveToTop() As Boolean)
    Dim Build() As Byte
    Dim X As Integer
    Dim Y As Integer
    ReDim Build(LBound(OriginalOrder) To UBound(OriginalOrder))
    For X = LBound(OriginalOrder) To UBound(OriginalOrder)
        If MoveToTop(X) Then
            Y = Y + 1
            Build(Y) = OriginalOrder(X)
        End If
    Next X
    For X = LBound(OriginalOrder) To UBound(OriginalOrder)
        If Not MoveToTop(X) Then
            Y = Y + 1
            Build(Y) = OriginalOrder(X)
        End If
    Next X
    OriginalOrder = Build
End Sub
Function RandomTarget(AtkPos As Byte) As Byte
    Dim X As Byte
    If ActNum = 2 Then
        X = OtherTeam(AtkPos)
    Else
        X = GetSmallRnd(1, 0)
        X = OtherTeam(TeamNum(AtkPos)) + X * 2
        If Current(X).HP <= 0 Then X = AllyNum(X)
    End If
    RandomTarget = X
End Function

'Calculate damages and stuff.  DoPreAttack contains the calculations that occur only once,
'regardless of the number of targets.
Sub DoPreAttack(ByVal AtkPos As Byte)
    Dim DefPos As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim F As Integer
    Dim M As Integer
    Dim RealMove As Integer
    Dim PreEffectTook As Integer
    Dim ConditionEffect As Byte
    Dim GoAhead As Boolean
    Dim TempMove As Move
    Dim TempPKMN() As Pokemon
    Dim CalledBy As Byte
    Dim TMove() As Integer
    Dim TrueAtkPos As Byte
    Dim Temp As String
    
    PreEffectTook = 1
    TrueAtkPos = AtkPos
    
    'This is mostly for Single Battles, it won't be used much for Doubles.
    'It's only temporery until the real target is decided.
    DefPos = OtherTeam(TeamNum(AtkPos))
    
    If BCondition(AtkPos).MoveUsed(1) <> 210 Then
        With BCondition(AtkPos)
            If .SemiInvul = nbFly Then
                M = 72
            ElseIf .SemiInvul = nbDig Then
                M = 43
            ElseIf .SemiInvul = nbBounce Then
                M = 261
            ElseIf .SemiInvul = nbDive Then
                M = 271
            ElseIf .Charging Then
                M = BCondition(AtkPos).MoveUsed(1) 'Yes, this IS necessary.  It prevents Encore/Metronome freakiness.
            ElseIf .EncoreMove <> 0 Then
                M = Current(AtkPos).Move(.EncoreMove)
            ElseIf .StuckMove <> 0 Then
                M = .StuckMove
            ElseIf .RepeatMove <> 0 And BattleMode = nbRBYBattle Then
                M = .RepeatMove
            Else
                M = BCondition(AtkPos).MoveUsed(1)
            End If
        End With
    Else
        M = BCondition(AtkPos).MoveUsed(1)
    End If
    
    TempMove = Moves(M)
    RealMove = M
    
    'Set the slot for Sketch
    For X = 1 To 4
        If Current(AtkPos).Move(X) = M Then Exit For
    Next X
    If X = 5 Then X = BCondition(AtkPos).LastMoveSlot
    If M = 210 Then X = 0
    BCondition(AtkPos).SketchSlot = X
        
    ReDim TempPKMN(ActNum)
    For X = 1 To ActNum
        TempPKMN(X) = Current(X)
    Next X
    
    'CalledBy Moves are set here.
    CalledBy = 0
    With TempMove
        Select Case .SpecialEffect
        'Metronome
        Case 69
            CalledBy = 1
            Select Case BattleMode
            Case nbRBYBattle
                Do
                    X = Int(Rnd * 251) + 1
                Loop While X = M Or X = 184 Or X = 210 Or X = 128 Or Not Moves(X).RBYMove
            Case nbGSCBattle
                Do
                    X = Int(Rnd * 251) + 1
                Loop While X = M Or X = 184 Or X = 210 Or X = 128
            Case nbAdvBattle
                Do
                    X = Int(Rnd * 354) + 1
                Loop While X = M Or X = 184 Or X = 210 Or X = 128 Or X = 256 Or X = 283 Or X = 311
            End Select
            M = X
            'M = GetMoveNum("Fly") 'DEBUG
            If Moves(M).Target = nbSelectedTarget Then BCondition(AtkPos).TargetNum = RandomTarget(AtkPos)
        'Mirror Move
        Case 73
            CalledBy = 2
            If BattleMode = nbAdvBattle Then
                M = BCondition(AtkPos).MirrorMove
            Else
                M = BCondition(OtherTeam(AtkPos)).LastMove
            End If
            Select Case M
            Case 0, 60, 180, 122, 124, 128, 190, 184
                M = 128
            Case 236
                If BattleMode <> nbRBYBattle Then M = 128
            End Select
            If BattleMode = nbGSCBattle Then
                For X = 1 To 4
                    If TempPKMN(AtkPos).Move(X) = M Then M = 128
                Next X
            End If
            If M = 128 Then CalledBy = 0
            If Moves(M).Target = nbSelectedTarget Then
                X = BCondition(AtkPos).MirrorPos
                If X = 0 Or X = AllyNum(AtkPos) Then X = RandomTarget(AtkPos)
                BCondition(AtkPos).TargetNum = X
            End If
        'Sleep Talk
        Case 100
            If TempPKMN(AtkPos).Condition <> 4 Then
                PreEffectTook = 0
            Else
                CalledBy = 3
                F = 0
                ReDim TMove(1 To 4)
                For X = 1 To 4
                    TMove(X) = 0
                    If TempPKMN(AtkPos).Move(X) = M Then TMove(X) = 1
                    If TempPKMN(AtkPos).Move(X) = 0 Then TMove(X) = 1
                    Select Case Moves(TempPKMN(AtkPos).Move(X)).SpecialEffect
                        Case 14, 32, 51, 69, 71, 73, 99, 100, 102, 122, 140, 154, 174
                            TMove(X) = 1
                    End Select
                    If TMove(X) = 1 Then F = F + 1
                Next X
                If F = 4 Then
                    CalledBy = 0
                    PreEffectTook = 0
                Else
                    If F = 3 Then
                        Y = 0
                    Else
                        Y = GetSmallRnd(3 - F)
                    End If
                    For X = 1 To 4
                        If TMove(X) = 0 Then
                            M = TempPKMN(AtkPos).Move(X)
                            If Y = 0 Then Exit For
                            Y = Y - 1
                        End If
                    Next X
                End If
                If Moves(M).Target = nbSelectedTarget Then BCondition(AtkPos).TargetNum = RandomTarget(AtkPos)
            End If
        'Assist
        Case 122
            CalledBy = 4
            ReDim TMove(0)
            F = TeamNum(AtkPos)
            For X = 1 To NumPoke
                If X <> Current(AtkPos).TeamNumber Then
                    For Y = 1 To 4
                        Z = PKMN(F, X).Move(Y)
                        If Not (Z = 0 Or Z = 122 Or Z = 184 Or Z = 210 Or Z = 128 Or Z = 256 Or Z = 283 Or Z = 284 Or Z = 311) Then
                            ReDim Preserve TMove(UBound(TMove) + 1)
                            TMove(UBound(TMove)) = Z
                        End If
                    Next Y
                End If
            Next X
            If UBound(TMove) = 0 Then
                CalledBy = 0
                PreEffectTook = 0
            Else
                M = TMove(GetSmallRnd(UBound(TMove), 1))
            End If
            If Moves(M).Target = nbSelectedTarget Then BCondition(AtkPos).TargetNum = RandomTarget(AtkPos)
        'Nature Power
        Case 154
            CalledBy = 5
            Select Case Terrain
            Case nbShortGrass, nbTallGrass
                M = 211 'Stun Spore
            Case nbVeryTallGrass
                M = 159 'Razor Leaf
            Case nbOcean
                M = 217 'Surf
            Case nbUnderwater
                M = 94 'Hydro Pump
            Case nbPond
                M = 24 'Bubblebeam
            Case nbsAnd
                M = 55 'Earthquake
            Case nbCave
                M = 181 'Shadow Ball
            Case nbMountain
                M = 167 'Rock Slide
            Case Else
                M = 221 'Swift
            End Select
            If Moves(M).Target = nbSelectedTarget Then BCondition(AtkPos).TargetNum = RandomTarget(AtkPos)
        End Select
    End With
    
    BCondition(AtkPos).MoveUsed(1) = M
    TempMove = ConvertMove(Moves(M), BattleMode)
    
    With TempMove

        'Moves that alter their base power
        Select Case .SpecialEffect
            'Bide
            Case 14
                If BCondition(AtkPos).BideCount = 0 Then
                    If BattleMode = nbAdvBattle Then
                        PreEffectTook = 2
                    Else
                        PreEffectTook = GetSmallRnd(3, 2)
                    End If
                End If
            'Counter
            Case 26
                If BattleMode = nbAdvBattle Then
                    If BCondition(AtkPos).LastPhDamage = 0 _
                    Or BCondition(AtkPos).LastPhHitter = 0 _
                    Or TeamNum(AtkPos) = TeamNum(BCondition(AtkPos).LastPhHitter) Then
                        PreEffectTook = 0
                    End If
                End If
            'Curse
            Case 28
                If TypeCheck(Current(AtkPos), nbGhost) Then .Target = nbSelectedTarget
            'Reversal/Flail
            Case 48
                .Power = 20
                If TempPKMN(AtkPos).HP <= TempPKMN(AtkPos).MaxHP * 0.688 Then .Power = 40
                If TempPKMN(AtkPos).HP <= TempPKMN(AtkPos).MaxHP * 0.354 Then .Power = 80
                If TempPKMN(AtkPos).HP <= TempPKMN(AtkPos).MaxHP * 0.208 Then .Power = 100
                If TempPKMN(AtkPos).HP <= TempPKMN(AtkPos).MaxHP * 0.104 Then .Power = 150
                If TempPKMN(AtkPos).HP <= TempPKMN(AtkPos).MaxHP * 0.042 Then .Power = 200
            'Frustration/Return
            Case 54
                'Note: Both of these will be given their max, since there's no Happiness tracking.
                .Power = 102
            'Fury Cutter
            Case 55
                BCondition(AtkPos).FuryCutter = BCondition(AtkPos).FuryCutter + 1
                If BCondition(AtkPos).FuryCutter = 6 Then BCondition(AtkPos).FuryCutter = 1
                .Power = 5
                For X = 1 To BCondition(AtkPos).FuryCutter
                    .Power = .Power * 2
                Next
            'Hidden Power
            Case 62
                If BattleMode = nbAdvBattle Then
                    TempMove.Power = HiddenPowerStrengthAdv(TempPKMN(AtkPos))
                    TempMove.Type = HiddenPowerTypeAdv(TempPKMN(AtkPos))
                Else
                    TempMove.Power = HiddenPowerStrength(TempPKMN(AtkPos).DV_Atk, TempPKMN(AtkPos).DV_Def, TempPKMN(AtkPos).DV_Spd, TempPKMN(AtkPos).DV_SAtk)
                    TempMove.Type = HiddenPowerType(TempPKMN(AtkPos).DV_Atk, TempPKMN(AtkPos).DV_Def)
                End If
            'Magnitude
            Case 66
                X = GetNextRnd
                Select Case X
                    Case Is < 14
                        .Power = 10
                        PreEffectTook = 4
                    Case Is < 39
                        .Power = 30
                        PreEffectTook = 5
                    Case Is < 90
                        .Power = 50
                        PreEffectTook = 6
                    Case Is < 167
                        .Power = 70
                        PreEffectTook = 7
                    Case Is < 218
                        .Power = 90
                        PreEffectTook = 8
                    Case Is < 243
                        .Power = 110
                        PreEffectTook = 9
                    Case Else
                        .Power = 150
                        PreEffectTook = 10
                End Select
            'RS Low Kick
            Case 70
                TempMove.Power = 1
            'Mirror Coat
            Case 72
                If BCondition(AtkPos).LastSpDamage = 0 _
                Or BCondition(AtkPos).LastSpHitter = 0 _
                Or TeamNum(AtkPos) = TeamNum(BCondition(AtkPos).LastSpHitter) Then
                    PreEffectTook = 0
                End If
            'RBY Petal Dance/Thrash
            Case 79
                If BattleMode = nbRBYBattle Then
                    If BCondition(AtkPos).StuckMove = 0 Then
                        PreEffectTook = GetSmallRnd(3, 2)
                    Else
                        PreEffectTook = GetSmallRnd(4, 1)
                    End If
                End If
            'Present
            Case 82
                X = GetSmallRnd(3)
                Select Case X
                Case 0: .Power = 40
                Case 1: .Power = 80
                Case 2: .Power = 120
                Case 3: .Power = 0
                End Select
            'Pursuit
            Case 86
                If BCondition(AtkPos).InPursuit Then PreEffectTook = 2
            'Solarbeam
            Case 102
                If CurrentWeather <> nbNoWeather And CurrentWeather <> nbSunny Then .Power = .Power \ 2
            'Rollout
            Case 121
                Select Case BCondition(AtkPos).Rollout
                    Case 0: PreEffectTook = 1 'Use PP
                    Case 4: PreEffectTook = 3 'Don't use PP, unlock the move
                    Case Else: PreEffectTook = 2 'Don't use PP
                End Select
                .Power = 30 * 2 ^ BCondition(AtkPos).Rollout
                If BCondition(AtkPos).DefenseCurl Then .Power = .Power * 2
            'Eruption/Water Spout
            Case 135
                .Power = Int((TempPKMN(AtkPos).HP / TempPKMN(AtkPos).MaxHP) * 150)
                If .Power = 0 Then .Power = 1
            'Facade
            Case 136
                If TempPKMN(AtkPos).Condition <> 1 And TempPKMN(AtkPos).Condition <> 4 Then .Power = .Power * 2
            'Focus Punch
            Case 140
                If BCondition(AtkPos).Damaged Then ConditionEffect = 15
            'Revenge
            Case 159
                If BCondition(AtkPos).Damaged Then .Power = .Power * 2
            'Secret Power
            Case 161
                .SpecialEffect = GetSPEffect
            'Spit Up
            Case 166
                .Power = .Power * BCondition(AtkPos).Stockpile
            'Weather Ball
            Case 178
                Select Case CurrentWeather
                Case nbRaining
                    .Type = nbWater
                Case nbSunny
                    .Type = nbFire
                Case nbSandstorm
                    .Type = nbRock
                Case nbHailstorm
                    .Type = nbIce
                End Select
                If CurrentWeather <> nbNoWeather Then .Power = .Power * 2
        End Select
        
        'Moves that hit both foes get half power
        If .Target = nbBothEnemies And ActNum = 4 Then
            Z = 0
            For Y = DefPos To ActNum Step 2
                If TempPKMN(Y).HP > 0 Then Z = Z + 1
            Next Y
            If Z = 2 Then .Power = .Power - (.Power \ 2) 'Round up
        End If
                        
        'Check if Spite drained the last of the PP before the move was used
        X = 0
        If BattleMode = nbRBYBattle Then X = 1
        If BCondition(AtkPos).Charging Or BCondition(AtkPos).Recharging Or BCondition(AtkPos).SemiInvul <> 0 Then X = 1
        If TempMove.SpecialEffect = 121 And PreEffectTook > 1 Then X = 1
        If BCondition(AtkPos).BideCount > 0 Then X = 1
        If TempMove.ID = 210 Then X = 1
        If CalledBy <> 0 Then X = 1
        If X = 0 Then
            For X = 1 To 4
                If TempPKMN(AtkPos).Move(X) = RealMove Then Exit For
            Next X
            If X <> 5 Then
                If Current(AtkPos).PP(X) = 0 Then ConditionEffect = 1
            End If
        End If
    
        'Adjust for minor conditions
        
        'Imprison
        For X = OtherTeam(TeamNum(AtkPos)) To ActNum Step 2
            If BCondition(X).Imprison Then
                For Y = 1 To 4
                    If BCondition(X).SealedMove(Y) = RealMove Then
                        ConditionEffect = 11
                        Exit For
                    End If
                Next Y
            End If
        Next X
    
        'Taunt
        If BCondition(AtkPos).TauntCount > 0 And TauntForbids(RealMove) Then ConditionEffect = 12
                
        'Disable
        If BCondition(AtkPos).DisabledMove > 0 Then
            If TempPKMN(AtkPos).Move(BCondition(AtkPos).DisabledMove) = RealMove Then ConditionEffect = 14
        End If
        
        'Flinch
        If BCondition(AtkPos).Flinching Then ConditionEffect = 18
                
        'Attract
        If BCondition(AtkPos).Attract <> 0 And (Not BCondition(AtkPos).InPursuit Or BattleMode = nbGSCBattle) Then
            If Odds(50) Then ConditionEffect = 9
        End If
        
        'Confusion
        If BCondition(AtkPos).ConfuseCounter > 0 And (Not BCondition(AtkPos).InPursuit Or BattleMode = nbGSCBattle) Then
            If Odds(50) Then
                X = GetAtk(AtkPos)
                Y = GetDef(AtkPos)
                PreEffectTook = Cap(Int((((2 * TempPKMN(AtkPos).Level / 5 + 2) * X * 40) / Y) / 50) + 2, Current(AtkPos).HP)
                If PreEffectTook = 0 Then PreEffectTook = 1
                ConditionEffect = 10
            End If
        End If
        
        'RBY Wrap
        If BattleMode = nbRBYBattle Then
            If BCondition(DefPos).RepeatMove > 0 Then ConditionEffect = 17
        End If
        
        'Truant
        If Current(AtkPos).Attribute = nbTruant And BCondition(AtkPos).Truant Then ConditionEffect = 2
        
        'Adjust for major conditions
        Select Case TempPKMN(AtkPos).Condition
            Case 4 'Asleep
                If CalledBy <> 3 And .SpecialEffect <> 100 And UproarCheck(AtkPos) = 0 And Not (.SpecialEffect = 101 And (CalledBy = 0 Or CalledBy = 3)) Then
                    ConditionEffect = 4
                End If
            Case 6 'Fully Paralysed
                If Not BCondition(AtkPos).InPursuit Or BattleMode = nbGSCBattle Then
                    If Odds(25) Then ConditionEffect = 6
                End If
            Case 7 'Thawing
                If BattleMode = nbRBYBattle Then
                    ConditionEffect = 7
                ElseIf M <> 69 And M <> 172 Then
                    If BattleMode = nbGSCBattle Then
                        If Not Odds(10) Then ConditionEffect = 7
                    Else
                        If Not Odds(25) Then ConditionEffect = 7
                    End If
                End If
        End Select
    End With
        
    'Okay, here's where we decide who gets whacked.
    'First off, if the game isn't a double battle, it's simple.
    'NOTE: Actually, the Double Battle targetting system
    '      must be used even if it's a Single in Ruby/Sapp.
    '      RS is just crazy that way.
    For X = 1 To ActNum
        BCondition(X).Targeted = False
    Next X
    If BattleMode <> nbAdvBattle Then
        BCondition(OtherTeam(AtkPos)).Targeted = True
    Else
        X = 0
        Y = OtherTeam(TeamNum(AtkPos))
        If BCondition(Y).FollowMe Then X = Y
        If BCondition(Y + 2).FollowMe Then X = Y + 2
        Y = X
        Select Case TempMove.Target
        Case nbSelectedTarget
            X = BCondition(AtkPos).TargetNum
            If X = 0 Then
                X = RandomTarget(AtkPos)
            ElseIf Current(X).HP <= 0 Then
                X = RandomTarget(AtkPos)
            End If
            If Y > 0 Then X = Y
            BCondition(X).Targeted = True
        Case nbBothEnemies
            For X = OtherTeam(TeamNum(AtkPos)) To ActNum Step 2
                If Current(X).HP > 0 Then BCondition(X).Targeted = True
            Next X
        Case nbEveryoneElse
            For X = 1 To ActNum
                If X <> AtkPos And Current(X).HP > 0 Then BCondition(X).Targeted = True
            Next X
        Case nbRandomEnemy
            X = RandomTarget(AtkPos)
            If Y > 0 Then X = Y
            BCondition(X).Targeted = True
        Case Else
            'Counter and Mirror Coat are exceptions
            If TempMove.SpecialEffect = 26 And PreEffectTook > 0 Then
                X = BCondition(AtkPos).LastPhHitter
                If X = 0 Then
                    BCondition(AtkPos).Targeted = True
                Else
                    If Y > 0 Then X = Y
                    BCondition(X).Targeted = True
                End If
            ElseIf TempMove.SpecialEffect = 72 And PreEffectTook > 0 Then
                X = BCondition(AtkPos).LastSpHitter
                If X = 0 Then
                    BCondition(AtkPos).Targeted = True
                Else
                    If Y > 0 Then X = Y
                    BCondition(X).Targeted = True
                End If
            Else
                BCondition(AtkPos).Targeted = True
            End If
        End Select
    End If
    
    'Now we start of the BattleString
    'IMPORTANT: Keep it in Binary until the end!
    Temp = ""
    For X = 1 To ActNum
        Temp = Temp & Bool2Bin(BCondition(X).Targeted)
    Next X
    BattleString = BattleString & "1" & Temp & Dec2Bin(AtkPos - 1, 2) & Dec2Bin(M, 9) & Dec2Bin(CalledBy, 3) & Dec2Bin(PreEffectTook, 10) & Dec2Bin(ConditionEffect, 5)
    
    'Kick off the turn
    Call DoPreTurn(AtkPos, M, CalledBy, PreEffectTook, ConditionEffect, GoAhead)
    If Not GoAhead Then GoTo EndProcess
    
    'Now do the hits
    Do
        For X = 1 To ActNum
            If BCondition(X).Targeted Then
                Call DoAttack(AtkPos, X, TempMove)
                BCondition(X).Targeted = False
                Exit For
            End If
        Next X
    Loop Until X = ActNum + 1
    
    'CalledBy moves cannot be Countered/Mirrored in GSC (They can in RS)
    If CalledBy > 0 And BattleMode <> nbAdvBattle Then
        BCondition(OtherTeam(AtkPos)).LastPhDamage = 0
        BCondition(OtherTeam(AtkPos)).LastSpDamage = 0
    End If
    
EndProcess:
    BCondition(TrueAtkPos).LastAttack = 0
    BCondition(TrueAtkPos).MirrorPos = 0
    Moved(TrueAtkPos) = False
    SwitchTo(TrueAtkPos) = 0
    Call SyncPKMN
End Sub

Private Sub DoAttack(ByRef AtkPos As Byte, ByVal DefPos As Byte, TempMove As Move)
    Dim TempPKMN() As Pokemon
    Dim ConditionEffect As Integer
    Dim M As Integer
    Dim X As Single
    Dim Y As Long
    Dim Z As Long
    Dim A As Byte
    Dim EffectTook As Integer
    Dim EffectParam As Integer
    Dim TrueAtkPos As Byte
    Dim TrueDefPos As Byte
    Dim AtkTeam As Byte
    Dim DefTeam As Byte
    Dim AtkAlly As Byte
    Dim DefAlly As Byte
    Dim HitCount As Integer
    Dim DamageTemp As Long
    Dim Damage() As Long
    Dim TempCHitDamage As Long
    Dim CHitOdds As Integer
    Dim CHitVar As Integer
    Dim CHit As Boolean
    Dim CHitArray() As Boolean
    Dim Hit As Boolean
    Dim AttackPower(1) As Integer
    Dim SAttackPower(1) As Integer
    Dim DefensePower(1) As Integer
    Dim SDefensePower(1) As Integer
    Dim ATK As Long
    Dim DEF As Long
    Dim Lev As Long
    Dim POW As Single
    Dim TMATCH As Single
    Dim STAB As Single
    Dim RAND As Integer
    Dim Ratio As Single
    Dim BaseType As Integer
    Dim F As Long
    Dim Temp As Integer
    Dim CHitTemp As Integer
    Dim TMove(1 To 4) As Integer
    Dim TempString As String
    Dim FocusBand As Boolean
    Dim TraitEffect() As Byte
    Dim TempArray() As Byte
    Dim KingsRock As Boolean
    
    On Error GoTo DoAttack_Error
    'Set the basic variables
    M = TempMove.ID
    TrueAtkPos = AtkPos
    TrueDefPos = DefPos
    EffectTook = 1
    ConditionEffect = 0
    HitCount = 1
    If AtkPos = DefPos Then TempMove.SelfMove = True
    ReDim Damage(1 To 1)
    ReDim CHitArray(1 To 1)
    ReDim TraitEffect(1 To 1)
    If BattleMode = nbRBYBattle Then
        BCondition(DefPos).LastPhDamage = 0
        BCondition(DefPos).LastPhHitter = 0
    End If
    
    'Magic Coat Freakiness
    If TempMove.MagicCoat And BCondition(DefPos).MagicCoat _
    And Not (TempMove.SoundMove And (Current(DefPos).Attribute = nbCacophony Or Current(DefPos).Attribute = nbSoundproof)) Then
        A = DefPos
        DefPos = AtkPos
        AtkPos = A
    End If
    
    If ActNum = 4 Then
        AtkTeam = TeamNum(AtkPos)
        DefTeam = TeamNum(DefPos)
        AtkAlly = AllyNum(AtkPos)
        DefAlly = AllyNum(DefPos)
    Else
        AtkTeam = AtkPos
        DefTeam = DefPos
        AtkAlly = 0
        DefAlly = 0
    End If
    ReDim TempPKMN(ActNum)
    For X = 1 To ActNum
        TempPKMN(X) = Current(X)
    Next X
        
    With TempMove
        'RBY Adjustment
        If BattleMode = nbRBYBattle Then
            Select Case TempPKMN(1).No
                'Magnemite/ton lose Steel
                Case 81, 82
                    TempPKMN(1).Type2 = nbNoType
            End Select
            Select Case TempPKMN(2).No
                'Magnemite/ton lose Steel
                Case 81, 82
                    TempPKMN(2).Type2 = nbNoType
            End Select
        End If
    
        'Set the type matchup
        TMATCH = TypeMatch(TempMove, DefPos, TempPKMN(DefPos).Type1, TempPKMN(DefPos).Type2)
    
        'Decided if the attack is a hit or miss.
        If .SpecialEffect = 44 Or .SpecialEffect = 120 Then
            'Vital Throw and Swift (+ clones) are unaffected by Acc or Evade
            Ratio = 1
        ElseIf (.SpecialEffect = 26 Or .SpecialEffect = 72) And BattleMode = nbGSCBattle Then
            'Counter and Mirror Coat are unaffected by Acc or Evade in GSC battles.
            Ratio = 1
        ElseIf BCondition(DefPos).Foresight Then
            'Foresight skips all this.
            Ratio = 1
        Else
            'Adjust for Accuracy mods
            Ratio = AccuracyMod(BCondition(AtkPos).AccuracyChange)
            'Adjust for CompoundEyes
            If TempPKMN(AtkPos).Attribute = nbCompoundEyes Then Ratio = Ratio * 1.3
            'Adjust for Hustle
            If TempPKMN(AtkPos).Attribute = nbHustle And AttackBase(.Type) = nbPysicalBased Then Ratio = Ratio * 0.8
            'Adjust for Evade mods
            'NOTE: Contrary to popular belief, Foresight itself IS affected by Evade mods.
            Ratio = Ratio * AccuracyMod(-BCondition(DefPos).EvadeChange)
            'Adjust for BrightPowder and LaxIncense
            If (TempPKMN(DefPos).Item = nbBrightPowder Or TempPKMN(DefPos).Item = nbLaxIncense) And .BrightPowder Then Ratio = Ratio * 0.922
            'Adjust for Sand Veil
            If TempPKMN(DefPos).Attribute = nbSandVeil And CurrentWeather = nbSandstorm Then Ratio = Ratio * 0.922
        End If
        
        Z = .Accuracy
        If M = 230 Then 'Adjust for Thunder's accuracy changes
            Select Case CurrentWeather
            Case nbSunny
                Z = 50
            Case nbRaining
                Z = 100
                Ratio = 1
            End Select
        ElseIf M = 18 Then 'Blizzard has perffect accuracy during Hail
            If CurrentWeather = nbHailstorm Then
                Z = 100
                Ratio = 1
            End If
        End If
        
        
        'And finally, calculate the hit
        Hit = Odds(Z * Ratio)
        
        'Self Affecting moves always hit
        If .Accuracy = 0 Or .SelfMove Then Hit = True
        
        'Mean Look Exception
        If .SpecialEffect = 67 Then Hit = True
        
        'Lock-On
        If BCondition(AtkPos).LockOn = DefPos Then Hit = True
        
        'RBY Wrap
        If BattleMode = nbRBYBattle Then
            If BCondition(AtkPos).RepeatCount > 0 Then Hit = True
        End If
            
        'Multihitter Hitcounts
        If .Power > 0 Then
            Select Case .SpecialEffect
                Case 10
                    If BattleMode = nbAdvBattle Then
                        X = GetNextRnd
                        Select Case X
                        Case Is < 96: HitCount = 2
                        Case Is < 192: HitCount = 3
                        Case Is < 224: HitCount = 4
                        Case Else: HitCount = 5
                        End Select
                    Else
                        HitCount = GetSmallRnd(5, 2)
                    End If
                Case 12
                    HitCount = 0
                    For X = 1 To NumPoke
                        If PKMN(AtkTeam, X).Condition = 1 Then
                            HitCount = HitCount + 1
                        End If
                    Next X
                Case 118
                    If BattleMode = nbAdvBattle Then
                        If BCondition(AtkPos).LockOn = DefPos Then
                            HitCount = 3
                        Else
                            HitCount = 1
                            If Odds(.Accuracy * Ratio) Then
                                HitCount = 2
                                If Odds(.Accuracy * Ratio) Then
                                    HitCount = 3
                                End If
                            End If
                        End If
                    Else
                        HitCount = GetSmallRnd(3, 1)
                    End If
                Case 19, 119
                    HitCount = 2
                Case Else
                    HitCount = 1
            End Select
        
            'Decide the damages
            If HitCount > 0 Then
                ReDim Damage(1 To HitCount)
                ReDim CHitArray(1 To HitCount)
            End If
            DamageTemp = 0
            For X = 1 To HitCount
                CHit = False
                'Calculate critical hit
                If BattleMode = nbRBYBattle Then
                    Y = IIf(.SpecialEffect = 4, 32, 4)
                    If BCondition(AtkPos).FocusEnergy Then
                        If StadiumMode Then
                            Y = Y * 4
                        ElseIf Current(AtkPos).BaseSpeed < Current(DefPos).BaseSpeed Then
                            Y = Y \ 4
                        End If
                    End If
                    Y = Int(TempPKMN(AtkPos).BaseSpeed * (Y / 8))
                    Y = Cap(Y, 255)
                    CHit = (GetNextRnd < Y)
                Else
                    CHitOdds = 0
                    'Focus Energy effect
                    If BCondition(AtkPos).FocusEnergy Then CHitOdds = CHitOdds + 1
                    'Good chance for Critical Hit effect
                    Select Case .SpecialEffect
                    Case 4, 124, 156: CHitOdds = CHitOdds + 2
                    Case 99: If BattleMode <> nbRBYBattle Then CHitOdds = CHitOdds + 2
                    End Select
                    'Scope Lens
                    If TempPKMN(AtkPos).Item = nbScopeLens Then CHitOdds = CHitOdds + 1
                    'Stick
                    If TempPKMN(AtkPos).No = 83 And TempPKMN(AtkPos).Item = nbStick Then CHitOdds = CHitOdds + 2
                    'Lucky Punch
                    If TempPKMN(AtkPos).No = 113 And TempPKMN(AtkPos).Item = nbLuckyPunch Then CHitOdds = CHitOdds + 2
                    'Lansat Berry
                    If BCondition(AtkPos).LansatBerry Then CHitOdds = CHitOdds + 1
                    CHitVar = GetNextRnd
                    If CHitOdds > 4 Then CHitOdds = 4
                    Select Case CHitOdds
                    Case 0
                        If CHitVar < 17 Then CHit = True
                    Case 1
                        If CHitVar < 32 Then CHit = True
                    Case 2
                        If CHitVar < 64 Then CHit = True
                    Case 3
                        If CHitVar < 85 Then CHit = True
                    Case 4
                        If CHitVar < 128 Then CHit = True
                    End Select
                End If
                
                If .SpecialEffect = 48 And BattleMode <> nbAdvBattle Then CHit = False
                If BattleMode = nbRBYBattle And X <> 1 Then CHit = CHitArray(1)
                If Current(DefPos).Attribute = nbBattleArmor Then CHit = False
                If Current(DefPos).Attribute = nbShellArmor Then CHit = False
                
                'Your variables for attack and defense
                AttackPower(0) = GetAtk(AtkPos)
                SAttackPower(0) = GetSAtk(AtkPos)
                DefensePower(0) = GetDef(DefPos)
                SDefensePower(0) = GetSDef(DefPos)
                If CHit Then
                    AttackPower(1) = GetAtk(AtkPos, False)
                    SAttackPower(1) = GetSAtk(AtkPos, False)
                    DefensePower(1) = GetDef(DefPos, False)
                    SDefensePower(1) = GetSDef(DefPos, False)
                    If BattleMode = nbAdvBattle Then
                        If AttackPower(1) < AttackPower(0) Then AttackPower(1) = AttackPower(0)
                        If SAttackPower(1) < SAttackPower(0) Then SAttackPower(1) = SAttackPower(0)
                        If DefensePower(1) > DefensePower(0) Then DefensePower(1) = DefensePower(0)
                        If SDefensePower(1) > SDefensePower(0) Then SDefensePower(1) = SDefensePower(0)
                    End If
                End If
                
                'Beat Up
                If .SpecialEffect = 12 Then
                    Z = 0
                    For Y = 1 To 6
                        If PKMN(AtkTeam, Y).Condition = 1 Then
                            Z = Z + 1
                        End If
                        If Z = X Then Exit For
                    Next Y
                    SAttackPower(0) = PKMN(AtkTeam, Y).BaseAttack
                    SAttackPower(1) = PKMN(AtkTeam, Y).BaseAttack
                    SDefensePower(0) = TempPKMN(DefPos).BaseDefense
                    SDefensePower(1) = TempPKMN(DefPos).BaseDefense
                End If
                
                '*****POWER BOOSTS/REDUCTION*****
                POW = .Power
                        
                'Ru/Sa Low Kick
                If .SpecialEffect = 70 Then
                    Select Case Current(DefPos).Weight
                    Case Is <= 100: POW = 20
                    Case Is <= 250: POW = 40
                    Case Is <= 500: POW = 60
                    Case Is <= 1000: POW = 80
                    Case Is <= 2000: POW = 100
                    Case Else: POW = 120
                    End Select
                End If
                  
                'Type Boosting items
                If .Type = 1 And M <> 210 And TempPKMN(AtkPos).Item = nbPinkBow Then POW = POW * 1.1
                If .Type = 1 And M <> 210 And TempPKMN(AtkPos).Item = nbPolkadotBow Then POW = POW * 1.1
                If .Type = 1 And M <> 210 And TempPKMN(AtkPos).Item = nbSilkScarf Then POW = POW * 1.1
                If .Type = 2 And TempPKMN(AtkPos).Item = nbCharcoal Then POW = POW * 1.1
                If .Type = 3 And TempPKMN(AtkPos).Item = nbMysticWater Then POW = POW * 1.1
                If .Type = 3 And TempPKMN(AtkPos).Item = nbSeaIncense Then POW = POW * 1.1
                If .Type = 4 And TempPKMN(AtkPos).Item = nbMagnet Then POW = POW * 1.1
                If .Type = 5 And TempPKMN(AtkPos).Item = nbMiracleSeed Then POW = POW * 1.1
                If .Type = 6 And TempPKMN(AtkPos).Item = nbNevermeltIce Then POW = POW * 1.1
                If .Type = 7 And TempPKMN(AtkPos).Item = nbBlackBelt Then POW = POW * 1.1
                If .Type = 8 And TempPKMN(AtkPos).Item = nbPoisonBarb Then POW = POW * 1.1
                If .Type = 9 And TempPKMN(AtkPos).Item = nbSoftSand Then POW = POW * 1.1
                If .Type = 10 And TempPKMN(AtkPos).Item = nbSharpBeak Then POW = POW * 1.1
                If .Type = 11 And TempPKMN(AtkPos).Item = nbTwistedSpoon Then POW = POW * 1.1
                If .Type = 12 And TempPKMN(AtkPos).Item = nbSilverPowder Then POW = POW * 1.1
                If .Type = 13 And TempPKMN(AtkPos).Item = nbHardStone Then POW = POW * 1.1
                If .Type = 14 And TempPKMN(AtkPos).Item = nbSpellTag Then POW = POW * 1.1
                If .Type = 15 And TempPKMN(AtkPos).Item = nbDragonFang Then POW = POW * 1.1
                If .Type = 16 And TempPKMN(AtkPos).Item = nbMetalCoat Then POW = POW * 1.1
                If .Type = 17 And TempPKMN(AtkPos).Item = nbBlackGlasses Then POW = POW * 1.1
                
                'For Triple Kick's increasing damage
                If .SpecialEffect = 118 Then POW = POW * X
                
                'Blaze, Overgrow, Torrent, and Swarm multiply their respective types by 1.5
                If Current(AtkPos).HP < Current(AtkPos).MaxHP / 3 Then
                    If Current(AtkPos).Attribute = nbBlaze And .Type = nbFire Then POW = POW * 1.5
                    If Current(AtkPos).Attribute = nbOvergrow And .Type = nbGrass Then POW = POW * 1.5
                    If Current(AtkPos).Attribute = nbTorrent And .Type = nbWater Then POW = POW * 1.5
                    If Current(AtkPos).Attribute = nbSwarm And .Type = nbBug Then POW = POW * 1.5
                End If
                
                RAND = Int(Rnd * 38) + 217
                'Flail/Reversal do not use the RAND variable
                If .SpecialEffect = 48 Then RAND = 255
                
                'For GSC Rage:
                If BCondition(AtkPos).RageCounter <> 0 And BattleMode = nbGSCBattle Then
                    RAND = RAND * Int(1 + (BCondition(AtkPos).RageCounter * 236 / 255))
                End If
                
                For Y = Abs(CInt(CHit)) To 0 Step -1
                    If CHit And Y = 0 Then
                        TempCHitDamage = DamageTemp
                        DamageTemp = 0
                    End If
                    
                    If AttackBase(.Type) = nbSpecialBased Then
                        ATK = SAttackPower(Y)
                        DEF = SDefensePower(Y)
                        If BattleMode = nbRBYBattle And Y = 0 And TeamBCondition(TeamNum(DefPos)).LightScreenCount <> 0 Then DEF = Cap(Rollover(SDefensePower(Y) * 2))
                    Else
                        ATK = AttackPower(Y)
                        DEF = DefensePower(Y)
                        If BattleMode = nbRBYBattle And Y = 0 And TeamBCondition(TeamNum(DefPos)).ReflectCount <> 0 Then DEF = Cap(Rollover(DefensePower(Y) * 2))
                    End If
                    
                    If .SpecialEffect = 42 And Not BattleMode = nbRBYBattle Then DEF = DEF / 2
                    
                    If CHit Then Lev = TempPKMN(AtkPos).Level * 2 Else Lev = TempPKMN(AtkPos).Level
                            
                    'Extra 50% for type matching
                    If TypeCheck(TempPKMN(AtkPos), .Type) Then STAB = 1.5 Else STAB = 1
                    
                    'Struggle and Beat Up do typeless damage
                    If M = 210 Or M = 13 Then
                        STAB = 1
                        TMATCH = 1
                    End If
                    
                    'Present
                    If .SpecialEffect = 82 And BattleMode = nbGSCBattle And Not (StadiumMode Or PresentRule) Then
                        ATK = 10
                        If TempPKMN(DefPos).Type2 = nbNoType Then
                            Lev = GetPresentValue(TempPKMN(DefPos).Type1)
                        Else
                            Lev = GetPresentValue(TempPKMN(DefPos).Type2)
                        End If
                        If TempPKMN(AtkPos).Type2 = nbNoType Then
                            DEF = GetPresentValue(TempPKMN(AtkPos).Type1)
                        Else
                            DEF = GetPresentValue(TempPKMN(AtkPos).Type2)
                        End If
                        If TMATCH = 0.5 Then TMATCH = 0.25
                    End If
                    
                    If DEF <= 0 Then DEF = 1
                    
                    'The Damage Formula
                    DamageTemp = Int(((((((((2 * Lev) \ 5 + 2) * ATK * POW) \ DEF) / 50) + 2) * STAB) * TMATCH) * RAND) \ 255
                    
                    'Light Screen and Reflect.  1/2 damage if it's protecting 1 poke, 2/3 if protecting 2 pokes.
                    Select Case .SpecialEffect
                    Case 26, 36, 47, 72, 76, 103, 109
                        'Do nothing here.
                    Case Else
                        If Not BattleMode = nbRBYBattle Then
                            Z = 0
                            If AttackBase(.Type) = nbSpecialBased Then
                                If TeamBCondition(TeamNum(DefPos)).LightScreenCount > 0 Then Z = 1
                            Else
                                If TeamBCondition(TeamNum(DefPos)).ReflectCount > 0 Then Z = 1
                            End If
                            If Z = 1 Then
                                If ActNum = 4 Then
                                    For Z = TeamNum(AtkPos) To 4 Step 2
                                        If TempPKMN(Y).HP <= 0 Then Exit For
                                    Next Z
                                    If Z > 4 Then
                                        DamageTemp = DamageTemp * 2 \ 3
                                    Else
                                        DamageTemp = DamageTemp \ 2
                                    End If
                                Else
                                    DamageTemp = DamageTemp \ 2
                                End If
                            End If
                        End If
                    End Select
                Next
                If CHit And (TempCHitDamage > DamageTemp Or BattleMode <> nbGSCBattle) Then DamageTemp = TempCHitDamage
                CHitArray(X) = CHit
                
                'DAMAGE BOOSTS
                
                'Weather boosts
                Select Case CurrentWeather
                Case nbRaining
                    If .Type = nbWater Then DamageTemp = DamageTemp * 1.5
                    If .Type = nbFire Then DamageTemp = DamageTemp / 2
                Case nbSunny
                    If .Type = nbFire Then DamageTemp = DamageTemp * 1.5
                    If .Type = nbWater Then DamageTemp = DamageTemp / 2
                End Select
                
                'Electric moves get 2x boost after Charge is used
                If .Type = nbElectr And BCondition(AtkPos).ChargeCount > 0 Then DamageTemp = DamageTemp * 2
                
                'Helping Hand gives a 1.5x boost
                If BCondition(AtkPos).HelpingHand Then DamageTemp = DamageTemp * 1.5
                
                'Smellingsalt gets a 2x boost if the target is paralyzed
                If .SpecialEffect = 164 And TempPKMN(DefPos).Condition = 6 Then DamageTemp = DamageTemp * 2
                
                'Thick Fat halves damage from Fire/Ice moves
                If Current(DefPos).Attribute = nbThickFat And (.Type = nbFire Or .Type = nbIce) Then DamageTemp = DamageTemp * 0.5
                
                'Hustle multiplies Physical Attacks by 1.5
                If Current(AtkPos).Attribute = nbHustle And AttackBase(.Type) = nbPysicalBased Then DamageTemp = DamageTemp * 1.5
                
                'Flash Fire multiplies Fire attacks by 1.5
                If Current(AtkPos).Attribute = nbFlashFire And BCondition(AtkPos).FlashFire And TempMove.Type = nbFire Then DamageTemp = DamageTemp * 1.5
                
                'Stomp/Minimize
                If BCondition(DefPos).Minimize And (M = 207 Or M = 257 Or M = 312 Or M = 277) Then DamageTemp = DamageTemp * 2
                
                'Water/Mud Sport halves the power of Fire/Electr moves
                If .Type = nbFire Then
                    If BCondition(AtkPos).WaterSport Then
                        DamageTemp = DamageTemp / 2
                    ElseIf ActNum = 4 Then
                        If BCondition(AllyNum(AtkPos)).WaterSport Then DamageTemp = DamageTemp / 2
                    End If
                ElseIf .Type = nbElectr Then
                    If BCondition(AtkPos).MudSport Then
                        DamageTemp = DamageTemp / 2
                    ElseIf ActNum = 4 Then
                        If BCondition(AllyNum(AtkPos)).MudSport Then DamageTemp = DamageTemp / 2
                    End If
                End If

                If DamageTemp = 0 And BattleMode <> nbRBYBattle Then DamageTemp = 1
                If TMATCH = 0 Then DamageTemp = 0
                If BattleMode = nbRBYBattle And DamageTemp = 0 And TMATCH > 0 Then
                    Hit = False
                    DamageTemp = 1
                End If
                Damage(X) = DamageTemp
                
                
                'This bit's for Multi-hit/Color Change freakiness....
                If HitCount > 1 And TempPKMN(DefPos).Attribute = nbColorChange Then
                    If BCondition(DefPos).Substitute > 0 Then
                        Z = 0
                        For Y = 1 To X - 1
                           Z = Z + Damage(X)
                        Next Y
                        If Z >= BCondition(DefPos).Substitute Then Y = 1 Else Y = 0
                    Else
                        Y = 1
                    End If
                    If Y = 1 Then
                         TempPKMN(DefPos).Type1 = TempMove.Type
                         TempPKMN(DefPos).Type2 = nbNoType
                         TMATCH = TypeMatch(TempMove, DefPos, TempPKMN(DefPos).Type1, TempPKMN(DefPos).Type2)
                    End If
                End If
            Next
            If HitCount > 1 Then CHit = False
            If HitCount = 0 Then
                HitCount = 1
                Damage(1) = 0
                EffectTook = 0
            End If
        End If
        
        Select Case .SpecialEffect
            'Bide
            Case 14
                Damage(1) = BCondition(AtkPos).BideDamage * 2
                If Damage(1) = 0 Then EffectTook = 0
            'Counter
            Case 26
                If GoesLastOK Then
                    If Moves(BCondition(DefPos).LastMove).SpecialEffect = 47 And BattleMode = nbGSCBattle Then
                        Damage(1) = TempPKMN(DefPos).HP
                    Else
                        Damage(1) = BCondition(AtkPos).LastPhDamage * 2
                    End If
                End If
                If Damage(1) = 0 Then EffectTook = 0
            'Dragon Rage
            Case 36
                Damage(1) = 40
            'OHKO
            Case 47
                Damage(1) = TempPKMN(DefPos).MaxHP
            'Mirror Coat
            Case 72
                If GoesLastOK Then
                    If Moves(BCondition(DefPos).LastMove).SpecialEffect = 47 And BattleMode <> nbAdvBattle Then
                        Damage(1) = TempPKMN(DefPos).HP
                    Else
                        Damage(1) = BCondition(AtkPos).LastSpDamage * 2
                    End If
                End If
                If Damage(1) = 0 Then EffectTook = 0
            'Night Shade
            Case 76
                Damage(1) = TempPKMN(AtkPos).Level
            'Psywave
            Case 85
                If BattleMode = nbAdvBattle Then
                    Damage(1) = Int((TempPKMN(AtkPos).Level / 100) * (GetSmallRnd(10, 0) * 10 + 50))
                Else
                    Damage(1) = Int(Rnd * (TempPKMN(AtkPos).Level * 1.5 - IIf(BattleMode = nbRBYBattle, 1, 0))) + 1
                End If
            'Pursuit
            Case 86
                If BCondition(AtkPos).InPursuit Then Damage(1) = Damage(1) * 2
            'SonicBoom
            Case 103
                 Damage(1) = 20
            'Super Fang
            Case 109
                'This does damage equal to half the
                'opponent's current HP, *ROUNDED UP*
                Damage(1) = TempPKMN(DefPos).HP \ 2 + TempPKMN(DefPos).HP Mod 2
            'Endeavor
            Case 134
                Damage(1) = TempPKMN(DefPos).HP - TempPKMN(AtkPos).HP
                If Damage(1) <= 0 Then
                    Damage(1) = 0
                    EffectTook = 0
                End If
        End Select
        
        'Insert code for special effects here
        If .SpecialPercent > 0 Then
            If TempPKMN(DefPos).Attribute = nbShieldDust Then
                EffectTook = 0
            Else
                Y = .SpecialPercent
                If TempPKMN(AtkPos).Attribute = nbSereneGrace Then Y = Y * 2
                If Y > 100 Then Y = 100
                If Not Odds(Y) Then EffectTook = 0
            End If
        End If
    
        If EffectTook = 1 Then
            'Work out details of certain moves
            Select Case .SpecialEffect
                'Poison, Paralysis, and Burns
                Case 18, 125, 39, 124, 81, 115, 119, 156
                    If Current(DefPos).Condition <> 1 Then EffectTook = 0
                'Attract
                Case 8
                    If Not AttractOK(AtkPos, DefPos) Or BCondition(DefPos).Attract Then
                        EffectTook = 0
                    End If
                'Baton Pass
                Case 11
                    Y = 0
                    For X = 1 To NumPoke
                        If PKMN(TeamNum(AtkPos), X).Condition = 8 Then Y = Y + 1
                    Next X
                    EffectTook = 0
                    If ActNum = 4 And Y < NumPoke - 2 Then EffectTook = 1
                    If ActNum = 2 And Y < NumPoke - 1 Then EffectTook = 1
                    SwitchTo(AtkPos) = 0
                'Wrap/Fire Spin/Etc.
                Case 15
                    EffectTook = GetSmallRnd(5, 2)
                    If BattleMode = nbRBYBattle Then
                        EffectTook = EffectTook - 1
                        If BCondition(AtkPos).StuckMove <> 0 Then Hit = True
                    End If
                'Flinch
                Case 16
                    'If Current(DefPos).Attribute = nbInnerFocus Then EffectTook = 0
                'Confusion
                Case 22, 110, 139
                    EffectTook = GetSmallRnd(4, 1)
                'Conversion
                Case 23
                    If BattleMode = nbRBYBattle Then
                        EffectParam = Current(DefPos).Type1
                        EffectParam = EffectParam + Current(DefPos).Type2 * 16
                    Else
                        A = 0
                        For X = 1 To 4
                            TMove(X) = 0
                            If TempPKMN(AtkPos).Move(X) > 0 And TempPKMN(AtkPos).Move(X) <> 38 Then
                                If Moves(TempPKMN(AtkPos).Move(X)).Type <> TempPKMN(AtkPos).Type1 Then
                                    A = A + 1
                                    TMove(X) = 1
                                End If
                            End If
                        Next X
                        If A = 0 Then
                            EffectTook = 0
                        Else
                            A = GetSmallRnd(A, 1)
                            For X = 1 To 4
                                If TMove(X) = 1 Then A = A - 1
                                If A = 0 Then Exit For
                            Next X
                            EffectParam = Moves(TempPKMN(AtkPos).Move(X)).Type
                        End If
                    End If
                'Conversion2
                Case 24
                    BaseType = BCondition(AtkPos).LastAttack
                    If BaseType = 0 Or BaseType = 38 Then
                        EffectTook = 0
                    Else
                        BaseType = Moves(BaseType).Type
                        ReDim TempArray(0)
                        For X = 1 To 17
                            If BattleMatrix(BaseType, X) < 1 And X <> TempPKMN(AtkPos).Type1 Then
                                ReDim Preserve TempArray(UBound(TempArray) + 1)
                                TempArray(UBound(TempArray)) = X
                            End If
                        Next X
                        If UBound(TempArray) = 0 Then
                            EffectTook = 0
                        Else
                            EffectParam = TempArray(GetSmallRnd(UBound(TempArray), 1))
                        End If
                    End If
                'Curse
                Case 28
                    If TypeCheck(Current(AtkPos), nbGhost) Then
                        If BattleMode = nbAdvBattle Then
                            If BCondition(DefPos).Curse Then EffectTook = 0
                        ElseIf BattleMode = nbGSCBattle Then
                            If BCondition(DefPos).Substitute > 0 Then EffectTook = 0
                        End If
                    End If
                'Detect
                Case 31
                    If GoesFirstOK And (BCondition(AtkPos).Substitute <= 0 Or BattleMode = nbAdvBattle) Then
                        If BCondition(AtkPos).ProtectPercent = 0 Then
                            BCondition(AtkPos).ProtectPercent = 100
                        Else
                            BCondition(AtkPos).ProtectPercent = BCondition(AtkPos).ProtectPercent / 2
                        End If
                        If Not Odds(BCondition(AtkPos).ProtectPercent) Then EffectTook = 0
                    Else
                        EffectTook = 0
                    End If
                'Disable
                Case 33
                    If BCondition(DefPos).DisabledMove = 0 Then
                        If BattleMode = nbRBYBattle Then
                            For X = 2 To 4
                                If Current(DefPos).Move(X) = 0 Then Exit For
                            Next X
                            EffectParam = GetSmallRnd(1, X - 1)
                            EffectTook = GetSmallRnd(4, 1)
                        Else
                            If BCondition(DefPos).LastMoveSlot = 0 Then
                                EffectTook = 0
                            Else
                                EffectParam = BCondition(DefPos).LastMoveSlot
                                EffectTook = GetSmallRnd(5, 2)
                            End If
                        End If
                    Else
                        EffectTook = 0
                    End If
                'Encore
                Case 40
                    X = BCondition(DefPos).LastMoveSlot
                    If X = 0 Then
                        EffectTook = 0
                    ElseIf Current(DefPos).PP(X) = 0 Then
                        EffectTook = 0
                    ElseIf BCondition(DefPos).Encore Then
                        EffectTook = 0
                    ElseIf BattleMode = nbGSCBattle And (Current(DefPos).Move(X) = 58 Or Current(DefPos).Move(X) = 128) Then
                        EffectTook = 0
                    Else
                        EffectTook = GetSmallRnd(6, 2)
                        EffectParam = X
                    End If
                'Endure
                Case 41
                    If (GoesFirstOK And BCondition(AtkPos).Substitute <= 0) Or BattleMode = nbAdvBattle Then
                        If BCondition(AtkPos).ProtectPercent = 0 Then
                            BCondition(AtkPos).ProtectPercent = 100
                        Else
                            BCondition(AtkPos).ProtectPercent = BCondition(AtkPos).ProtectPercent / 2
                        End If
                        If Not Odds(BCondition(AtkPos).ProtectPercent) Then EffectTook = 0
                    Else
                        EffectTook = 0
                    End If
                'Leech Seed
                Case 46
                    If TypeCheck(TempPKMN(DefPos), nbGrass) Or BCondition(DefPos).LeechSeed > 0 Then EffectTook = 0
                'OHKO
                Case 47
                    If (BattleMode = nbRBYBattle And GetSpd(AtkPos) < GetSpd(DefPos)) _
                    Or (BattleMode <> nbRBYBattle And Current(AtkPos).Level < Current(DefPos).Level) Then
                        EffectTook = 0
                        Damage(1) = 0
                    End If
                'Focus Energy
                Case 52
                    If BCondition(AtkPos).FocusEnergy = True Then EffectTook = 0
                'Future Sight
                Case 56
                    Damage(1) = 0
                    If BCondition(DefPos).FutureSight.count = 0 Then
                        EffectParam = GetNextRnd
                        Hit = True
                    Else
                        EffectTook = 0
                    End If
                'Light Screen
                Case 57
                    If TeamBCondition(TeamNum(AtkPos)).LightScreenCount > 0 Then EffectTook = 0
                'Hidden Power
                Case 62
                    'EffectTook = .Type
                'Sleep
                Case 64
                    If TempPKMN(DefPos).Condition <> 1 Or RuleCheck(DefTeam, 1) Then
                        EffectTook = 0
                    Else
                        EffectTook = GetSmallRnd(8, 2)
                        If BattleMode = nbRBYBattle Then EffectTook = EffectTook - 1
                    End If
                'Magnitude
                Case 66
                    'EffectTook = Magnitude
                'Mimic
                Case 71
                    If BattleMode = nbRBYBattle Then
                        Z = 0
                        For X = 1 To 4
                            TMove(X) = 0
                            Temp = 0
                            For Y = 1 To 4
                                If Current(DefPos).Move(X) = Current(AtkPos).Move(Y) Then Exit For
                            Next Y
                            If Y = 5 And Current(DefPos).Move(X) <> 0 Then Z = Z + 1: TMove(X) = 1
                        Next X
                        If Z = 0 Then
                            EffectTook = 0
                        Else
                            Do
                                EffectTook = GetSmallRnd(4, 1)
                            Loop Until TMove(EffectTook) = 1
                            EffectParam = TempPKMN(DefPos).Move(EffectTook)
                        End If
                    Else
                        X = BCondition(DefPos).LastMoveSlot
                        If X = 0 Then
                            EffectTook = 0
                        Else
                            EffectParam = Current(DefPos).Move(BCondition(DefPos).LastMoveSlot)
                            For X = 1 To 4
                                If TempPKMN(AtkPos).Move(X) = EffectParam Then EffectTook = 0
                            Next X
                        End If
                    End If
                'Mirror Move
                Case 73
                    EffectTook = 0
                    Damage(1) = 0
                'Mist
                Case 74
                    If BCondition(AtkPos).Mist = True Or TeamBCondition(AtkTeam).MistCount > 0 Then EffectTook = 0
                'Nightmare
                Case 77
                    If BCondition(DefPos).Nightmare = True Then EffectTook = 0
                    If Current(DefPos).Condition <> 4 Then EffectTook = 0
                'Outrage/Thrash
                Case 79
                    EffectTook = GetSmallRnd(3, 2) 'Move Length
                    EffectParam = GetSmallRnd(4, 1) 'Confusion Duration
                'Present
                Case 82
                    If .Power = 0 Then EffectTook = 2
                'Psych Up
                Case 84
                    If BCondition(DefPos).AttackChange = 0 _
                    And BCondition(DefPos).DefenseChange = 0 _
                    And BCondition(DefPos).SpeedChange = 0 _
                    And BCondition(DefPos).SAttackChange = 0 _
                    And BCondition(DefPos).SDefenseChange = 0 _
                    And BCondition(DefPos).EvadeChange = 0 _
                    And BCondition(DefPos).AccuracyChange = 0 _
                    And BattleMode = nbGSCBattle Then
                        EffectTook = 0
                    End If
                 'Rain Dance
                Case 88
                    If BattleMode = nbAdvBattle And CurrentWeather = nbRaining Then EffectTook = 0
                'Reflect
                Case 91
                    If TeamBCondition(TeamNum(AtkPos)).ReflectCount > 0 Then EffectTook = 0
                'Roar/Whirlwind
                Case 93
                    F = 0
                    For X = 1 To NumPoke
                        If PKMN(TeamNum(DefPos), X).Condition = 8 Then F = F + 1
                    Next X
                    If F < NumPoke - ActNum / 2 Then
                        X = TempPKMN(DefPos).TeamNumber
                        If ActNum = 4 Then Y = TempPKMN(AllyNum(DefPos)).TeamNumber Else Y = 0
                        While X = TempPKMN(DefPos).TeamNumber Or X = Y Or PKMN(TeamNum(DefPos), X).Condition = 8
                            X = Int(Rnd * NumPoke) + 1
                        Wend
                        EffectTook = X
                    Else
                        EffectTook = 0
                    End If
                    If Not GoesLastOK Or BattleMode = nbRBYBattle Then EffectTook = 0
                    If BattleMode = nbAdvBattle And Current(AtkPos).Level < Current(DefPos).Level Then
                        If Odds(25) Then EffectTook = 0
                    End If
                    EffectParam = GetSmallRnd(1) 'For Trace
                'Safeguard
                Case 94
                    If TeamBCondition(AtkTeam).SafeguardCount > 0 Then EffectTook = 0
                'Sandstorm
                Case 95
                    If CurrentWeather = nbSandstorm Then EffectTook = 0
                'Sketch
                Case 97
                    If BattleMode = nbAdvBattle Then
                        EffectParam = BCondition(DefPos).SketchSlot
                        If EffectParam = 0 Then
                            EffectTook = 0
                        Else
                            EffectParam = Current(DefPos).Move(EffectParam)
                        End If
                        If BCondition(AtkPos).TransformedTo > 0 Then EffectTook = 0
                    Else
                        EffectParam = BCondition(DefPos).LastMove
                        If EffectParam = 0 Then EffectTook = 0
                        If BCondition(DefPos).TransformedTo > 0 Then EffectTook = 0
                    End If
                    For X = 1 To 4
                        If TempPKMN(AtkPos).Move(X) = EffectParam Then EffectTook = 0
                    Next X
                'Snore
                Case 101
                    If TempPKMN(AtkPos).Condition <> 4 Or Current(DefPos).Attribute = nbInnerFocus Then
                        EffectTook = 0
                    End If
                'Spikes
                Case 104
                    If (BattleMode = nbAdvBattle And TeamBCondition(OtherTeam(AtkTeam)).Spikes = 3) _
                    Or (BattleMode = nbGSCBattle And TeamBCondition(OtherTeam(AtkTeam)).Spikes = 1) Then
                        EffectTook = 0
                    End If
                'Spite
                Case 105
                    If BCondition(DefPos).LastMoveSlot = 0 Then
                        EffectTook = 0
                    Else
                        EffectTook = Cap(GetSmallRnd(5, 2), CInt(TempPKMN(DefPos).PP(BCondition(DefPos).LastMoveSlot)))
                        EffectParam = BCondition(DefPos).LastMoveSlot
                    End If
                    'If EffectTook = 0 And BattleMode = nbGSCBattle Then EffectTook = 6
                'Thief, Recycle, and Covet
                Case 114, 157
                    EffectParam = GetItemEffect
                'Sunny Day
                Case 108
                    If BattleMode = nbAdvBattle And CurrentWeather = nbSunny Then EffectTook = 0
                'Transform
                Case 116
                    If BattleMode <> nbRBYBattle And (BCondition(AtkPos).TransformedTo > 0 Or BCondition(DefPos).TransformedTo > 0) Then EffectTook = 0
                'Tri Attack
                Case 117
                    If TempPKMN(DefPos).Condition = 7 Then
                        EffectTook = 4
                    Else
                        EffectTook = GetSmallRnd(3, 1)
                    End If
               'Fake Out
                Case 137
                    If BCondition(AtkPos).FakeOut = 0 Then Damage(1) = 0
                    If Current(DefPos).Attribute = nbInnerFocus Or Damage(1) = 0 Then EffectTook = 0
                'Hail
                Case 143
                    If CurrentWeather = nbHailstorm Then EffectTook = 0
                'Helping Hand
                Case 144
                    If ActNum = 2 Then
                        EffectTook = 0
                    ElseIf Current(AtkAlly).HP <= 0 Then
                        EffectTook = 0
                    End If
                'Imprison
                Case 145
                    A = 0
                    For Y = 1 To 4
                        For F = DefTeam To ActNum Step 2
                            For Z = 1 To 4
                                If TempPKMN(AtkPos).Move(Y) = TempPKMN(F).Move(Z) Then
                                    A = 1
                                    Exit For
                                End If
                            Next Z
                            If A Then Exit For
                        Next F
                        If A Then Exit For
                    Next Y
                    If A = 0 Then EffectTook = 0
                'Refresh
                Case 158
                    If TempPKMN(AtkPos).Condition = 1 Or TempPKMN(AtkPos).Condition = 4 Then EffectTook = 0
                'Role Play
                Case 160
                    If Current(DefPos).Attribute = nbWonderGuard Then EffectTook = 0
                'Skill Swap
                Case 162
                    If Current(AtkPos).Attribute = nbWonderGuard Or Current(DefPos).Attribute = nbWonderGuard Then EffectTook = 0
                'Trick
                Case 173
                    If Current(AtkPos).Item = nbNoItem And Current(DefPos).Item = nbNoItem Then
                        EffectTook = 0
                    End If
                'Uproar
                Case 174
                    If BCondition(AtkPos).StuckMove = 0 Then EffectTook = GetSmallRnd(5, 2)
                'Wish
                Case 179
                    If BCondition(AtkPos).WishCount > 0 Then EffectTook = 0
                'Yawn
                Case 180
                    If UproarCheck(DefPos) <> 0 _
                    Or Current(DefPos).Condition > 1 _
                    Or BCondition(DefPos).YawnCount > 0 _
                    Or RuleCheck(DefTeam, 1) Then
                        EffectTook = 0
                    Else
                        EffectTook = GetSmallRnd(8, 2)
                    End If
            End Select
        End If
            

        If Not BattleMode = nbRBYBattle Then
            If Not .SelfMove Then
            
                'Compensate for Flying
                If BCondition(DefPos).SemiInvul = nbFly Then
                    Select Case M
                    Case 79, 230, 247, 332
                        'Thunder, Whirlwind, and Sky Uppercut are unaffected by Fly
                    Case 85, 240
                        '2x damage for Gust and Twister
                        Damage(1) = Damage(1) * 2
                    Case Else
                        'Everything else misses
                        Hit = False
                    End Select
                End If
                
                'Compensate for Digging
                If BCondition(DefPos).SemiInvul = nbDig And Not .SelfMove Then
                    Select Case M
                    Case 67, 79
                        'Fissure is unaffected by Dig
                    Case 55, 114
                        '2x damage for Earthquake and Magnitude
                        Damage(1) = Damage(1) * 2
                    Case Else
                        'Everything else misses
                        Hit = False
                    End Select
                End If
                
                'Compensate for Diving
                If BCondition(DefPos).SemiInvul = nbDive Then
                    Select Case M
                    Case 217
                        '2x damage for Surf
                        Damage(1) = Damage(1) * 2
                    End Select
                End If
                
                'Compensate for Bouncing
                If BCondition(DefPos).SemiInvul = nbBounce Or BCondition(DefPos).SemiInvul = nbDive Then
                    'Everything misses.  (I think... look into this.)
                    Hit = False
                End If

            End If
        Else
            If BCondition(DefPos).SemiInvul <> 0 And Not .SelfMove And .SpecialEffect <> 44 Then
                Hit = False
            End If
        End If
        
        'Attacks that only work while opponent is asleep
        If (.SpecialEffect = 37 Or .SpecialEffect = 77) And TempPKMN(DefPos).Condition <> 4 Then Hit = False
    
        'King's Rock
        If .KingsRock = True _
        And Hit = True _
        And Damage(1) > 0 And Current(DefPos).Attribute <> nbInnerFocus Then
            KingsRock = Odds(11.7)
        End If
        
        'Focus Band
        If TempPKMN(DefPos).Item = nbFocusBand Then FocusBand = Odds(11.7)
                 
        'Contact Move Traits
        ReDim TraitEffect(1 To HitCount)
        If .PhysMove And Hit Then
            For X = 1 To HitCount
                If Damage(X) > 0 Then
                    Select Case TempPKMN(DefPos).Attribute
                    Case nbStatic, nbFlameBody, nbPoisonPoint, nbCuteCharm
                        If Odds(30) Then TraitEffect(X) = 1
                    Case nbEffectSpore
                        If Odds(10) Then
                            TraitEffect(X) = GetSmallRnd(3, 1)
                            If TraitEffect(X) = 3 Then TraitEffect(X) = GetSmallRnd(9, 3)
                        End If
                    Case nbRoughSkin
                        TraitEffect(X) = 1
                    End Select
                End If
            Next X
        End If
        
        'Damage Adjust
        For X = 1 To HitCount
            Damage(X) = Cap(Damage(X), 1023)
        Next X
        If HitCount = 1 And Damage(1) = 0 And BattleMode <> nbRBYBattle Then HitCount = 0: Erase Damage
    End With

    BattleString = BattleString & Bool2Bin(Hit) & Dec2Bin(HitCount, 3) & Dec2Bin(EffectTook, 4) & Dec2Bin(EffectParam, 10) & Bool2Bin(FocusBand)
    For X = 1 To HitCount
        BattleString = BattleString & Dec2Bin(Damage(X), 10) & Bool2Bin(CHitArray(X)) & Dec2Bin(TraitEffect(X), 4)
    Next X
    
    Call DoTurn(TrueAtkPos, TrueDefPos, M, Hit, EffectTook, EffectParam, FocusBand, Damage, CHitArray, TraitEffect, KingsRock)

    Exit Sub

DoAttack_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DoAttack of Class Module BattleData"
    If InVBMode Then
        Stop
        Resume
    End If
End Sub
'Start off the turn
Sub DoPreTurn(AtkPos As Byte, ByVal M As Integer, ByVal CalledBy As Integer, ByVal PreEffectTook As Long, ByVal ConditionEffect As Byte, ByRef GoAhead As Boolean)
'             ^^^---- For Snatch/MagicCoat purposes, AtkPos is NOT ByVal.
    Dim X As Byte
    Dim Y As Integer
    Dim Z As Byte
    Dim DoPP As Boolean
    Dim TempMove As Move
    Dim TrueAtkPos As Byte
    Dim RealMove As Integer
    Dim SkipMessage As Boolean
    Dim SkipLMS As Boolean
    Dim Blocked() As Boolean
    Dim s As Boolean
    
    ReDim Blocked(1 To ActNum)
    GoAhead = True
    SkipMessage = False
    TrueAtkPos = AtkPos
    TempMove = ConvertMove(Moves(M), BattleMode)
    
'    If BMessStyle = 1 Then
'        Call AddMessage(FText(72, PName(AtkPos), Current(AtkPos).Nickname), , , , True)
'    End If

    Select Case CalledBy
    Case 1: CalledBy = 122
    Case 2: CalledBy = 128
    Case 3: CalledBy = 190
    Case 4: CalledBy = 256
    Case 5: CalledBy = 311
    End Select
    RealMove = IIf(CalledBy > 0, CalledBy, M)

    DoPP = True
    BCondition(AtkPos).DestinyBond = False
    BCondition(AtkPos).Grudge = False
    If BCondition(AtkPos).FakeOut > 0 Then
        BCondition(AtkPos).FakeOut = BCondition(AtkPos).FakeOut - 1
    End If
    
    'Everything in this block takes place before the XXX used XXX line
    
    'Reset the Fury Cutter counter if the move isn't Fury Cutter
    If TempMove.SpecialEffect <> 55 Then
        BCondition(AtkPos).FuryCutter = 0
    End If
    'Reset the Rage counter if the move isn't Rage
    If TempMove.SpecialEffect <> 87 Then
        BCondition(AtkPos).RageCounter = 0
    End If
    'Reset the Rollout counter if the move isn't Rollout
    If TempMove.SpecialEffect <> 121 Then
        BCondition(AtkPos).Rollout = 0
    End If
    'Reset the Protect/Endure counter if the move isn't either of those.
    If TempMove.SpecialEffect <> 31 And TempMove.SpecialEffect <> 41 Then
        BCondition(AtkPos).ProtectPercent = 0
    End If
            
    'Uproar Awakening
    If UproarCheck(AtkPos) <> 0 And Current(AtkPos).Condition = 4 Then
        Call AddMessage(223, nbActive, AtkPos)
        Current(AtkPos).Condition = 1
        Current(AtkPos).ConditionCount = 0
        If CalledBy > 0 Then
            If Moves(CalledBy).SpecialEffect = 100 Then
                M = CalledBy
                CalledBy = 0
                TempMove = Moves(M)
            End If
        End If
    End If
    
    'Truant
    If ConditionEffect = 2 Then
        Call AddMessage(293, nbActive, AtkPos)
        BCondition(AtkPos).SemiInvul = 0
        BCondition(AtkPos).Charging = False
        BCondition(AtkPos).Recharging = False
        BCondition(AtkPos).ProtectPercent = 0
        BCondition(AtkPos).Rollout = 0
        BCondition(AtkPos).FuryCutter = 0
        BCondition(AtkPos).MoveUsed(1) = 0
        Call ClearStuckMove(AtkPos)
        GoAhead = False
        DoPP = False
        GoTo EndProcess
    End If
    
    'Recharging
    If BCondition(AtkPos).Recharging Then
        Call AddMessage(77, nbActive, AtkPos)
        DoPP = False
        BCondition(AtkPos).Recharging = False
        GoAhead = False
        GoTo EndProcess
    End If
    
    'Random thawing
    If Current(AtkPos).Condition = 7 And ConditionEffect <> 7 Then
        If M = 69 Or M = 172 Then
            If BattleMode = nbAdvBattle Then
                Call AddMessage(186, nbActive, AtkPos, nbMove, M)
                Current(AtkPos).Condition = 1
            End If
        Else
            If BattleMode = nbAdvBattle Then
                Call AddMessage(185, nbActive, AtkPos)
                Current(AtkPos).Condition = 1
            Else
                Unfreezing(AtkPos) = True
                ConditionEffect = 7
            End If
        End If
    End If
    
    'RBY Awakening
    If Current(AtkPos).Condition = 4 And BattleMode = nbRBYBattle Then
        Current(AtkPos).ConditionCount = Current(AtkPos).ConditionCount - 1
        If Current(AtkPos).ConditionCount = 0 Then
            Current(AtkPos).Condition = 1
            Current(AtkPos).Resting = False
            Call AddMessage(173, nbActive, AtkPos)
        End If
    End If
    
    'Pursuit
    If TempMove.SpecialEffect = 86 And PreEffectTook = 2 Then
        BCondition(AtkPos).InPursuit = True
        For X = 1 To ActNum
            If BCondition(X).Targeted Then Exit For
        Next X
        If Not BCondition(X).BeingPursued And BattleMode = nbAdvBattle Then
            BCondition(X).BeingPursued = True
            Call AddMessage(297, nbPlayer, TeamNum(X), nbActive, X)
        End If
    End If

    'Sleep Message
    'This still happens even with Sleep Talk and Snore, which
    'is why it's removed from the following Select Case
    If Current(AtkPos).Condition = 4 Then Call AddMessage(54, nbActive, AtkPos)

    'Immobilization due to a condition.  Decided in DoPreAttack, enacted here.
    GoAhead = False
    Select Case ConditionEffect
    Case 4 'Fast Asleep
        'Do Nothing
    Case 6 'Fully Paralysed
        Call AddMessage(55, nbActive, AtkPos)
    Case 7 'Frozen Solid
        If Current(AtkPos).Condition = 7 Then Call AddMessage(56, nbActive, AtkPos)
    Case 11 'Move failed due to Imprison
        Call AddMessage(53, nbActive, AtkPos, nbMove, RealMove)
    Case 12 'Move failed due to Taunt
        Call AddMessage(52, nbActive, AtkPos, nbMove, RealMove)
    Case 14 'Move failed due to Disable
        Call AddMessage(50, nbActive, AtkPos, nbMove, Current(AtkPos).Move(BCondition(AtkPos).DisabledMove))
    Case 17 'RBY Wrapped
        Call AddMessage(40, nbActive, AtkPos)
    Case 18 'Flinched
        Call AddMessage(47, nbActive, AtkPos)
    Case Else
        GoAhead = True
    End Select
    
    'Confused/Attracted message
    If BCondition(AtkPos).Confuse And GoAhead And Not BCondition(AtkPos).InPursuit Then
        If BCondition(AtkPos).ConfuseCounter = 0 And BattleMode = nbAdvBattle Then
            BCondition(AtkPos).Confuse = False
            Call AddMessage(192, nbActive, AtkPos)
        Else
            If BCondition(AtkPos).ConfuseCounter < 5 And BattleMode = nbAdvBattle Then
                BCondition(AtkPos).ConfuseCounter = BCondition(AtkPos).ConfuseCounter - 1
            End If
            Call AddMessage(191, nbActive, AtkPos)
            If ConditionEffect = 10 Then 'Hurt itself in its Confusion
                Call AddMessage(58, nbActive, AtkPos)
                Current(AtkPos).HP = Current(AtkPos).HP - PreEffectTook
                GoAhead = False
            End If
        End If
    End If
    
    If BCondition(AtkPos).Attract <> 0 And GoAhead And Not BCondition(AtkPos).InPursuit Then
        Call AddMessage(197, nbActive, AtkPos, nbActive, BCondition(AtkPos).Attract)
        If ConditionEffect = 9 Then 'Immobilized by love
            Call AddMessage(57, nbActive, AtkPos)
            GoAhead = False
        End If
    End If
    
    If Not GoAhead Then
        BCondition(AtkPos).SemiInvul = 0
        BCondition(AtkPos).Charging = False
        BCondition(AtkPos).Recharging = False
        BCondition(AtkPos).StuckCount = 0
        BCondition(AtkPos).ProtectPercent = 0
        BCondition(AtkPos).Rollout = 0
        BCondition(AtkPos).FuryCutter = 0
        BCondition(AtkPos).MoveUsed(1) = 0
        Call ClearStuckMove(AtkPos)
        DoPP = False
        GoTo EndProcess
    End If
    
    '[Name] used [move] message
    If CalledBy > 0 Then Call AddMessage(4, nbActive, AtkPos, nbMove, CalledBy)
    If CalledBy = 311 Then Call AddMessage(260, nbMove, 311, nbMove, M)
    
    'Call DoAnim(BattleWindow, nbaMove + CalledBy, AtkPos, 0)
    
    If ConditionEffect = 15 Then 'Focus Punch failure
        Call AddMessage(284, nbActive, AtkPos)
        GoAhead = False
        GoTo EndProcess
    End If
    
    Select Case TempMove.SpecialEffect
    'Bide
    Case 14
        If BattleMode <> nbAdvBattle And BCondition(AtkPos).Substitute > 0 Then
            PreEffectTook = 0
        Else
            If BCondition(AtkPos).BideCount = 0 Then
                BCondition(AtkPos).BideCount = PreEffectTook
                BCondition(AtkPos).BideDamage = 0
                BCondition(AtkPos).LastAnyHitter = 0
            Else
                BCondition(AtkPos).BideCount = BCondition(AtkPos).BideCount - 1
                If BCondition(AtkPos).BideCount > 0 Then Call AddMessage(75, nbActive, AtkPos)
                DoPP = False
                SkipMessage = True
            End If
            If BCondition(AtkPos).BideCount > 0 Then
                GoAhead = False
            Else
                For X = 1 To ActNum
                    BCondition(X).Targeted = False
                Next X
                Call AddMessage(76, nbActive, AtkPos)
                SkipMessage = True
                X = BCondition(AtkPos).LastAnyHitter
                If X = 0 Then
                    BCondition(AtkPos).Targeted = True
                Else
                    BCondition(X).Targeted = True
                End If
            End If
        End If
    'Wrap
    Case 15
        If BattleMode = nbRBYBattle Then
            If BCondition(AtkPos).RepeatCount > 0 Then
                Call AddMessage(295, nbActive, AtkPos)
                SkipMessage = True
                DoPP = False
            End If
        End If
    'Dig, Fly, Bounce, and Dive
    Case 32, 51, 125, 132
        If BCondition(AtkPos).SemiInvul = 0 Then
            Select Case TempMove.SpecialEffect
            Case 32
                BCondition(AtkPos).SemiInvul = nbDig
                Call AddMessage(72, nbActive, AtkPos)
            Case 51
                BCondition(AtkPos).SemiInvul = nbFly
                Call AddMessage(71, nbActive, AtkPos)
            Case 125
                BCondition(AtkPos).SemiInvul = nbBounce
                Call AddMessage(74, nbActive, AtkPos)
            Case 132
                BCondition(AtkPos).SemiInvul = nbDive
                Call AddMessage(73, nbActive, AtkPos)
            End Select
            GoAhead = False
            SkipMessage = True
            If BattleMode = nbRBYBattle Then DoPP = False
        Else
            If BattleMode <> nbRBYBattle Then DoPP = False
            BCondition(AtkPos).SemiInvul = 0
            SkipLMS = True
        End If
    'Skull Bash
    Case 98
        If BCondition(AtkPos).Charging = False Then
            BCondition(AtkPos).Charging = True
            Call AddMessage(69, nbActive, AtkPos)
            If Not BattleMode = nbRBYBattle Then Call ModifyPKMN(AtkPos, 2, 1)
            GoAhead = False
            SkipMessage = True
            If BattleMode = nbRBYBattle Then DoPP = False
        Else
            If BattleMode <> nbRBYBattle Then DoPP = False
            BCondition(AtkPos).Charging = False
            SkipLMS = True
        End If
    'Sky Attack
    Case 99
        If BCondition(AtkPos).Charging = False Then
            BCondition(AtkPos).Charging = True
            Call AddMessage(70, nbActive, AtkPos)
            GoAhead = False
            SkipMessage = True
            If BattleMode = nbRBYBattle Then DoPP = False
        Else
            If BattleMode <> nbRBYBattle Then DoPP = False
            BCondition(AtkPos).Charging = False
            SkipLMS = True
        End If
    'Solarbeam
    Case 102
        If BCondition(AtkPos).Charging Then
            BCondition(AtkPos).Charging = False
            If BattleMode <> nbRBYBattle Then DoPP = False
            SkipLMS = True
        ElseIf CurrentWeather <> nbSunny Then
            If BattleMode = nbRBYBattle Then DoPP = False
            BCondition(AtkPos).Charging = True
            Call AddMessage(68, nbActive, AtkPos)
            GoAhead = False
            SkipMessage = True
        End If
    End Select
    
    If Not SkipMessage Then Call AddMessage(4, nbActive, AtkPos, nbMove, M)
    
    'Outta PP
    'Yes, the [Name] used [move] message SHOULD show up
    'even though it's not happening.  Go figure.
    If ConditionEffect = 1 Then
        Call AddMessage(62)
        BCondition(AtkPos).MoveUsed(1) = 0
        BCondition(AtkPos).StuckCount = 0
        BCondition(AtkPos).ProtectPercent = 0
        BCondition(AtkPos).Rollout = 0
        BCondition(AtkPos).FuryCutter = 0
        Call ClearStuckMove(AtkPos)
        GoAhead = False
        GoTo EndProcess
    End If
    
    'Self-Thaw
    If (M = 69 Or M = 172) And Current(AtkPos).Condition = 7 Then
        Call AddMessage(185, nbActive, AtkPos)
        Current(AtkPos).Condition = 1
    End If
    
    'Special Stuff: Anything with irregular PP depletion or that happens
    '               regardless of Protect/Endure or any other protection
    '               status is taken care of here.
    Select Case TempMove.SpecialEffect
    'SelfDestruct
    Case 42
        For X = 1 To ActNum
            If Current(X).Attribute = nbDamp And Current(X).HP > 0 Then Exit For
        Next X
        If X = ActNum + 1 Then
            BCondition(AtkPos).ExplosionHP = Current(AtkPos).HP
            Current(AtkPos).HP = 0
            SelfLoss(TeamNum(AtkPos)) = True
            BCondition(AtkPos).FaintLast = True
        Else
            GoAhead = False
            Call AddMessage(135, nbActive, X, nbTrait, nbDamp, nbActive, AtkPos, nbMove, M)
        End If
    'Magnitude
    Case 66
        If PreEffectTook <> 0 Then Call AddMessage(213, nbNumber, PreEffectTook)
    'Mirror Move
    Case 73
        Call AddMessage(59)
        GoAhead = False
    'RBY Petal Dance/Thrash
    Case 79
        If BattleMode = nbRBYBattle Then
            If BCondition(AtkPos).StuckMove = 0 Then
                BCondition(AtkPos).StuckMove = M
                BCondition(AtkPos).StuckCount = PreEffectTook
            Else
                BCondition(AtkPos).StuckCount = BCondition(AtkPos).StuckCount - 1
                If BCondition(AtkPos).StuckCount = 0 Then
                    BCondition(AtkPos).StuckMove = 0
                    If BCondition(AtkPos).Confuse = False Then
                        BCondition(AtkPos).Confuse = True
                        BCondition(AtkPos).ConfuseCounter = PreEffectTook
                        'Call AddMessage(195, nbActive, AtkPos)
                    End If
                End If
            End If
        End If
    'Rage (RBY and Adv)
    Case 87
        If BattleMode <> nbGSCBattle Then BCondition(AtkPos).RageCounter = 1
        If BattleMode = nbRBYBattle And BCondition(AtkPos).StuckMove = 156 Then DoPP = False
    'Sleep Talk
    Case 100
        'It would be in CalledBy if it actually worked.
        PreEffectTook = 0
    'Snore
    Case 101
        If Current(AtkPos).Condition <> 4 Then PreEffectTook = 0
    'Rollout
    Case 121
        BCondition(AtkPos).Rollout = BCondition(AtkPos).Rollout + 1
        If BCondition(AtkPos).Rollout = 6 Then BCondition(AtkPos).Rollout = 1
        Select Case PreEffectTook
        Case 1
            BCondition(AtkPos).StuckMove = M
        Case 2
            DoPP = False
        Case 3
            DoPP = False
            BCondition(AtkPos).StuckMove = 0
        End Select
    'Memento
    Case 149
        Current(AtkPos).HP = 0
        SelfLoss(AtkPos) = True
        BCondition(AtkPos).FaintLast = True
    'Spit Up
    Case 166
        If BCondition(AtkPos).Stockpile = 0 Then
            Call AddMessage(60)
            GoAhead = False
        Else
            BCondition(AtkPos).Stockpile = 0
        End If
    End Select
    
    'The Move Failed message
    If PreEffectTook = 0 Then
        Call AddMessage(38, nbActive, AtkPos)
        BCondition(AtkPos).ProtectPercent = 0
        GoAhead = False
        GoTo EndProcess
    End If
    
    'For LightningRod
    If TempMove.Type = 4 And TempMove.Target = nbSelectedTarget And ActNum = 4 Then
        Y = 0
        For X = OtherTeam(TeamNum(AtkPos)) To ActNum Step 2
            If Current(X).HP > 0 And Current(X).Attribute = nbLightningRod Then Y = X '<-- No Exit For
        Next X
        If Y > 0 Then
            If Not BCondition(Y).Targeted Then
                Call AddMessage(162, nbActive, Y, nbTrait, nbLightningRod)
                For Z = 1 To ActNum
                    BCondition(Z).Targeted = False
                Next Z
                BCondition(Y).Targeted = True
            End If
        End If
    End If
                
    'For Snatch
    X = 0
    If TempMove.Target = nbSelfAffecting Or TempMove.Target = nbTeamAffecting Then X = 1
    Select Case TempMove.SpecialEffect
    Case 11, 23, 24, 28, 31, 41, 113, 144, 145, 152, 157, 176, 179
        X = 0
    Case 84
        X = 1
    End Select
    If X = 1 Then
        For X = 1 To ActNum
            For Y = 1 To ActNum
                If BCondition(Y).Snatch = X Then Exit For
            Next Y
            If Y <= ActNum Then
                Call AddMessage(279, nbActive, Y, nbActive, AtkPos)
                If TempMove.SpecialEffect = 84 Then
                    For Z = 1 To ActNum
                        BCondition(Z).Targeted = False
                    Next Z
                    BCondition(AtkPos).Targeted = True
                End If
                AtkPos = Y
                BCondition(Y).Snatch = 0
                SkipLMS = True
            End If
        Next X
    End If
    
    'Soundproof is really weird.
    Z = 0
    If TempMove.SoundMove Then
        For X = 1 To ActNum
            If Current(X).Attribute = nbCacophony Or Current(X).Attribute = nbSoundproof Then
                If BCondition(X).Targeted Then
                    Blocked(X) = True
                    Z = 1
                End If
            End If
        Next X
    End If
    
    'For Magic Coat
    If TempMove.MagicCoat Then
        For X = 1 To ActNum
            If BCondition(X).Targeted And Not Blocked(X) And BCondition(X).MagicCoat Then
                For Y = X To ActNum
                    If BCondition(Y).Targeted Then Blocked(Y) = True
                Next Y
            End If
        Next X
    End If
    Y = 0
    For X = 1 To ActNum
        If BCondition(X).Targeted And Not Blocked(X) Then Y = Y + 1
    Next X
    If Y = 0 Then SkipLMS = True
    
EndProcess:
    With BCondition(TrueAtkPos)
        
        'Set last used move for attacks such as Disable, Encore, and Spite
        If .MoveUsed(1) = 0 Then M = 0: CalledBy = 0
        If M <> 0 Or BattleMode <> nbRBYBattle Then .LastMove = M
        'If BattleMode = nbRBYBattle And Not DoPP Then .LastMove = 0
        If CalledBy <> 0 Then M = CalledBy
        If .LastMove = 0 Or s Then
            .LastMoveSlot = 0
        ElseIf (DoPP Or BattleMode = nbRBYBattle) And Not SkipLMS Then 'See note
            For X = 1 To 4
                If Current(TrueAtkPos).Move(X) = M Then Exit For
            Next X
            If X = 5 Then .LastMoveSlot = 0 Else .LastMoveSlot = X
        End If
        'NOTE: The DoPP check above is nessisary for a weird situation involving
        'CalledBy moves.  Say you have a Pokemon with the moves Metronome and
        'Dig.  You use Metronome and it just happens to call Dig.  Well, on
        'the first turn, CalledBy = Metronome, but on the second turn, NetBattle
        'is going to think that the REAL Dig (not Metronome) has been used.  To
        'prevent this, .LastMoveSlot is not touched when DoPP is false.  This
        'keeps the last .LastMoveSlot in there without being reset.  This
        'effectively fixes the problem for every 2-turn move, but it might cause
        'trouble later on if DoPP needs be False but .LastMoveSlot still needs to
        'be set.  Worse comes to worse, a seperate variable can be used.
        
        'Deplete PP
        If DoPP Then
            If .LastMoveSlot = 0 Then M = 0
            For X = 1 To 4
                If Current(TrueAtkPos).Move(X) = M Then Exit For
            Next X
            If X <> 5 And M <> 0 Then
                'Pressure
                Y = -1
                For Z = 1 To ActNum
                    If BCondition(Z).Targeted And Current(Z).Attribute = nbPressure And Z <> AtkPos Then Y = Y - 1
                Next Z
                Call DoPPChange(TrueAtkPos, X, Y)
            End If
        End If
        
        'Mimic, Sketch, Struggle, and Transform are weird
        If (CalledBy <> 0 And BattleMode = nbGSCBattle) Then .LastMove = 0: .LastMoveSlot = 0
        If .LastMove = 124 Or .LastMove = 184 Or .LastMove = 210 Then .LastMove = 0: .LastMoveSlot = 0
        If .LastMove = 236 And BattleMode <> nbRBYBattle Then .LastMove = 0: .LastMoveSlot = 0
        
        'Set MultiTarget for the Miss Message
        Z = 0
        For X = 1 To ActNum
           If BCondition(X).Targeted Then Z = Z + 1
        Next X
        MultiTarget = (Z > 1)
    
        'Set the stuck Choice Band move
        If Current(TrueAtkPos).Item = nbChoiceBand And .LastMoveSlot <> 0 Then .ChoiceBandSlot = .LastMoveSlot
    End With
End Sub

'Execute a single hit.  Note that the order of everything has been tested
'to be exact to the gameboy.
Sub DoTurn(ByRef AtkPos As Byte, ByVal DefPos As Byte, ByVal M As Integer, _
ByVal Hit As Boolean, ByVal EffectTook As Long, ByVal EffectParam As Long, _
ByVal FocusBand As Boolean, Damage() As Long, CHitArray() As Boolean, _
TraitEffect() As Byte, Optional ByVal KingsRock As Boolean)
    Dim Temp As Single
    Dim Temp2 As Integer
    Dim TempMove As Move
    Dim X As Long
    Dim Y As Long
    Dim Z As Integer
    Dim A As Integer
    Dim B As Integer
    Dim CHit As Boolean
    Dim HitSubstitute As Boolean
    Dim DoPP As Boolean
    Dim TempString As String
    Dim TempString2 As String
    Dim AtkTeam As Byte
    Dim DefTeam As Byte
    Dim AtkAlly As Byte
    Dim DefAlly As Byte
    Dim TMATCH As Single
    'Dim TypeMessage As Byte
    Dim SkipTypeMessage As Boolean
    Dim Damaged As Boolean
    Dim HitCount As Byte
    Dim Multihitter As Boolean
    Dim MultihitExit As Boolean
    Dim MultihitDamage As Integer
    
    'Pokémon Ruby/Sapphire Turn Order Testing Results
    '
    'Failure Order
    '
    'Blocked by Soundproof/Cacophony
    'Blocked by Protect/Detect
    'Blocked by Substitute
    'Blocked by Uproar (Sleep)
    'Blocked by Levitate/Wonder Guard
    'Blocked by Status Ailment Trait/Suction Cups/Ingrain
    '[Pokemon] already has a [ailment]
    'It Doesn't Affect [Pokemon] (Non-Damaging Attacks)
    'Attack Failed
    'Attack Missed
    'Blocked or Absorbed by Water/Volt Absorb
    'It Doesn't Affect [Pokemon] (Damaging Attacks)
    'Flash Fire Activation/Block
    'Blocked by Safeguard
    '----Conversion2 and Mirror Move set here----
    'Blocked by Mist
    'Blocked by Clear Body/Hyper Cutter/Keen Eye
    '
    'Damage Order
    '
    'Single Hitter:
    'Damage Infliction
    'NetBattle's Damage message
    'Critical Hit Message
    'SE/NVE Message
    'Endure Check
    'False Swipe Check
    'Focus Band Check
    'Move Effect Execution
    'Attacker Faint Check
    'Defender Faint Check
    'Rage Build
    'Color Change
    'Contact Traits
    'Rough Skin Damage
    'Attacker Faint Check
    'Shell Bell
    '
    '
    'Multi Hitter:
    '--Loop Start--
    'Damage Infliction
    'Critical Hit message
    'Endure Check
    'False Swipe Check
    'Focus Band Check
    'Rage Build
    'Color Change
    'Contact Traits
    'Rough Skin Damage
    'Attacker Faint Check
    'Shell Bell
    '---Loop End---
    'SE/NVE Message
    'Hit Count Message
    'NetBattle's Damage message
    'Move Effect Execution
    'Attacker Faint Check
    'Defender Faint Check

    
    
    'First, set a few preliminary variables
    HitCount = 0
    On Error Resume Next
    HitCount = UBound(Damage)
    On Error GoTo DoTurnError
    If HitCount = 0 Then HitCount = 1: ReDim Damage(1 To 1)
    TempMove = ConvertMove(Moves(M), BattleMode)
    
    'Magic Coat
    If TempMove.MagicCoat And BCondition(DefPos).MagicCoat _
    And Not (TempMove.SoundMove And (Current(DefPos).Attribute = nbCacophony Or Current(DefPos).Attribute = nbSoundproof)) Then
        BCondition(DefPos).MagicCoat = False
        Call AddMessage(277, nbActive, AtkPos, nbMove, M, nbMove, 302)
        A = DefPos
        DefPos = AtkPos
        AtkPos = A
        If ActNum = 4 And TempMove.Target = nbBothEnemies Then
            If (BCondition(AllyNum(AtkPos)).Targeted Or Current(AllyNum(AtkPos)).HP <= 0) And Current(AllyNum(DefPos)).HP > 0 Then
                BCondition(AllyNum(DefPos)).Targeted = True
            End If
            For X = AtkPos To 4 Step 2
                BCondition(X).Targeted = False
            Next X
        End If
    End If
    
    If ActNum = 4 Then
        AtkTeam = TeamNum(AtkPos)
        DefTeam = TeamNum(DefPos)
        AtkAlly = AllyNum(AtkPos)
        DefAlly = AllyNum(DefPos)
    Else
        AtkTeam = AtkPos
        DefTeam = DefPos
        AtkAlly = 0
        DefAlly = 0
    End If
    Select Case M
    Case 15, 34, 50, 67, 68, 75, 84, 93, 112, 114, 127, 133, 153, 164, 165, 179, 198, 215, 275, 327, 148
        TempMove.Power = 1
'    Case 148
'        If EffectTook <> 2 Then TempMove.Power = 1
'    Case 336
'        If BCondition(AtkPos).Stockpile = 0 Then TempMove.Power = 0
    End Select
    Select Case TempMove.SpecialEffect
    Case 56
        'For now, Future Sight is powerless and perfectly accurate
        TempMove.Power = 0
        Hit = True
    Case 62
        If BattleMode = nbAdvBattle Then
            TempMove.Power = HiddenPowerStrengthAdv(Current(AtkPos))
            TempMove.Type = HiddenPowerTypeAdv(Current(AtkPos))
        Else
            TempMove.Power = HiddenPowerStrength(Current(AtkPos).DV_Atk, Current(AtkPos).DV_Def, Current(AtkPos).DV_Spd, Current(AtkPos).DV_SAtk)
            TempMove.Type = HiddenPowerType(Current(AtkPos).DV_Atk, Current(AtkPos).DV_Def)
        End If
    Case 161
        TempMove.SpecialEffect = GetSPEffect
    Case 10, 12, 19, 118, 119
        Multihitter = True
    'Weather Ball
    Case 178
        Select Case CurrentWeather
        Case nbRaining: TempMove.Type = nbWater
        Case nbSunny: TempMove.Type = nbFire
        Case nbSandstorm: TempMove.Type = nbRock
        Case nbHailstorm: TempMove.Type = nbIce
        End Select
    End Select
    
    TMATCH = TypeMatch(TempMove, DefPos)
    
    'Decide what messages to display
    'NOTE: Now that damaging moves with 0 power are set to a power of 1
    '      at the beginning of the sub, this is partially unnecessary.
'    TypeMessage = 0
'    Select Case TempMove.SpecialEffect
'    Case 48, 54, 55, 62, 66 'These moves get both Super/Not Very Effective messages AND Doesn't Affect message
'        TypeMessage = 2
'    Case 14, 26, 36, 45, 72, 76, 85, 103, 109, 134 'These moves ONLY get the Doesn't Affect message.
'        TypeMessage = 1
'    Case 47 'OHKOs are different in RBY
'        TypeMessage = IIf(BattleMode = nbRBYBattle, 2, 1)
'    Case 82
'        If EffectTook <> 2 Then TypeMessage = 2 'Present is freaky.
'    End Select
    
    Select Case TempMove.SpecialEffect
    Case 14, 26, 36, 45, 72, 76, 85, 103, 109, 134 'These moves ONLY get the Doesn't Affect message.
        SkipTypeMessage = True
    Case 47 'OHKOs are different in RBY
        SkipTypeMessage = Not (BattleMode = nbRBYBattle)
    Case Else
        SkipTypeMessage = False
    End Select
    
    'First let's make sure the damage will actually happen...
    '---FAILURE CHECKS---
    
    'SOUNDPROOF/CACOPHONY
    If TempMove.SoundMove And Not TempMove.SelfMove Then
        If Current(DefPos).Attribute = nbSoundproof Or Current(DefPos).Attribute = nbCacophony Then
            Call AddMessage(155, nbActive, DefPos, nbTrait, Current(DefPos).Attribute, nbMove, M)
            EffectTook = 0
            If BCondition(AtkPos).Soundproofed = 0 Then BCondition(AtkPos).Soundproofed = 2
            GoTo EndProcess
        Else
            BCondition(AtkPos).Soundproofed = 1
        End If
    End If
    
    'PROTECT/DETECT
    If BCondition(DefPos).Protected = True And Not TempMove.SelfMove Then
        Call AddMessage(41, nbActive, DefPos)
        EffectTook = 0
        BCondition(AtkPos).Charging = False
        BCondition(AtkPos).FuryCutter = 0
        BCondition(AtkPos).Recharging = False
        BCondition(AtkPos).Rollout = 0
        BCondition(AtkPos).SemiInvul = 0
        If BCondition(AtkPos).StuckMove > 0 Then
            If Moves(BCondition(AtkPos).StuckMove).SpecialEffect = 121 Then
                BCondition(AtkPos).StuckMove = 0
            End If
        End If
        If BattleMode <> nbAdvBattle Then BCondition(AtkPos).RageCounter = 0
        GoTo EndProcess
    End If
    
    'Brick Break activates right here
    If TempMove.SpecialEffect = 126 And Hit Then
        'The freaky thing about this is that it removes LS/Refl. from
        'the opponent even if you're attacking your ally... not only that,
        'the barriers are still busted even if you're attacking a Ghost type...
        X = OtherTeam(TeamNum(AtkPos))
        Y = 0
        If TeamBCondition(X).ReflectCount <> 0 Then
            TeamBCondition(X).ReflectCount = 0
            Y = 1
        End If
        If TeamBCondition(X).LightScreenCount <> 0 Then
            TeamBCondition(X).LightScreenCount = 0
            Y = 1
        End If
        If Y = 1 Then Call AddMessage(286)
    End If

    'SUBSTITUTE
    If BCondition(DefPos).Substitute > 0 Then
        Select Case TempMove.SpecialEffect
        'Life leeching moves and Forsight fail in GSC only
        Case 1, 37, 78, 53
            If BattleMode = nbGSCBattle Then
                Call AddMessage(165, nbActive, DefPos, nbMove, 213)
                EffectTook = 0
                GoTo EndProcess
            End If
        'Poison and Sleep work in RBY
        Case 18, 64, 119
            If BattleMode <> nbRBYBattle Then
                EffectTook = 0
                If TempMove.Power = 0 Then
                    Call AddMessage(165, nbActive, DefPos, nbMove, 213)
                    GoTo EndProcess
                End If
            End If
        'Swagger, Flatter, and Mean Look fail in RS
        Case 110, 139, 67
            If BattleMode = nbAdvBattle Then
                Call AddMessage(165, nbActive, DefPos, nbMove, 213)
                EffectTook = 0
                GoTo EndProcess
            End If
        'Various effects fail in all versions
        Case 2, 9, 15 To 17, 20 To 22, 25, 27, 39, 46, 50, 65, 77, 78, 81, 96, 97, 101, 111, 114, 115, 117, 124, 125, 137, 138, 147, 150, 151, 156, 180
            EffectTook = 0
            If TempMove.Power = 0 Then
                Call AddMessage(165, nbActive, DefPos, nbMove, 213)
                GoTo EndProcess
            End If
        'And a few exceptions
        Case 28
            If TypeCheck(Current(AtkPos), nbGhost) Then
                EffectTook = 0
                Call AddMessage(165, nbActive, DefPos, nbMove, 213)
                GoTo EndProcess
            End If
        Case 149
            EffectTook = 0
            Call AddMessage(48)
            GoTo EndProcess
        End Select
    End If
                
    'UPROAR
    If TempMove.SpecialEffect = 64 Then
        If UproarCheck(DefPos) = DefPos Then
            Call AddMessage(227, nbActive, DefPos)
            EffectTook = 0
            GoTo EndProcess
        ElseIf UproarCheck(DefPos) <> 0 Then
            Call AddMessage(229, nbActive, DefPos)
            EffectTook = 0
            GoTo EndProcess
        End If
    End If

    'DAMAGE BLOCKING TRAITS
    If TempMove.Power > 0 Then
        Y = 0
        If Current(DefPos).Attribute = nbWonderGuard And TMATCH <= 1 And M <> 210 And M <> 133 Then
            Call AddMessage(42, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
            Y = 1
        ElseIf Current(DefPos).Attribute = nbLevitate And TempMove.Type = 9 Then
            Call AddMessage(43, nbActive, DefPos, nbType, 9, nbTrait, Current(DefPos).Attribute)
            Y = 1
        ElseIf Current(DefPos).Attribute = nbSturdy And TempMove.SpecialEffect = 47 Then
            Call AddMessage(134, nbActive, DefPos, nbTrait, nbSturdy)
            Y = 1
        End If
        If Y = 1 Then
            EffectTook = 0
            BCondition(AtkPos).Charging = False
            BCondition(AtkPos).FuryCutter = 0
            BCondition(AtkPos).Recharging = False
            BCondition(AtkPos).Rollout = 0
            BCondition(AtkPos).SemiInvul = 0
            If BattleMode <> nbAdvBattle Then BCondition(AtkPos).RageCounter = 0
            If BCondition(AtkPos).StuckMove > 0 Then
                If Moves(BCondition(AtkPos).StuckMove).SpecialEffect = 121 Then
                    BCondition(AtkPos).StuckMove = 0
                End If
            End If
            GoTo EndProcess
        End If
    End If
    
    'STATUS AILMENT PREVENTION TRAITS
    X = 0
    Select Case TempMove.SpecialEffect
    'Flinch
    Case 16, 101, 137
        If Current(DefPos).Attribute = nbInnerFocus Then
            If Moves(M).SpecialPercent = 0 Or Moves(M).SpecialPercent = 100 Then
                X = 147
            Else
                X = -1
            End If
        End If
    'Attract
    Case 8: If Current(DefPos).Attribute = nbOblivious Then X = 146
    'Paralyze
    Case 18, 125: If Current(DefPos).Attribute = nbLimber Then X = 142
    'Confusion
    Case 22: If Current(DefPos).Attribute = nbOwnTempo Then X = 145
    'Burn
    Case 39, 124: If Current(DefPos).Attribute = nbWaterVeil Then X = 143
    'Sleep
    Case 64, 180: If Current(DefPos).Attribute = nbInsomnia Or Current(DefPos).Attribute = nbVitalSpirit Then X = 230
    'Poison
    Case 81, 115, 119, 156: If Current(DefPos).Attribute = nbImmunity Then X = 144
    'Roar/Whirlwind
    Case 93
        If Current(DefPos).Attribute = nbSuctionCups Then
            X = 151
        ElseIf BCondition(DefPos).Ingrain Then
            X = 269
        End If
    End Select
    If X <> 0 Then
        EffectTook = 0
        If TempMove.Power = 0 Then
            If X > 0 Then Call AddMessage(X, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
            GoTo EndProcess
        End If
    End If

    'PRE-EXISTING AILMENT
    X = 0
    Select Case TempMove.SpecialEffect
    Case 81, 115, 119, 156: If Current(DefPos).Condition = 2 Or Current(DefPos).Condition = 3 Then X = 177
    Case 64: If Current(DefPos).Condition = 4 Then X = 171
    Case 39, 124: If Current(DefPos).Condition = 5 Then X = 182
    Case 18, 125: If Current(DefPos).Condition = 6 Then X = 189
    Case 22: If BCondition(DefPos).Confuse Then X = 194
    End Select
    If X > 0 Then
        EffectTook = 0
        If TempMove.Power = 0 Then
            Call AddMessage(X, nbActive, DefPos)
            GoTo EndProcess
        End If
    End If
    
    'TYPE IMMUNITY (NON-DAMAGING)
    Select Case TempMove.SpecialEffect
    'Ice-types can't freeze
    Case 17
        If TempMove.Type = 6 And TypeCheck(Current(DefPos), nbIce) Then
            EffectTook = 0
            If TempMove.Power = 0 Then
                Call AddMessage(45, nbActive, DefPos)
                GoTo EndProcess
            End If
        End If
    'Fire-types can't burn
    Case 39
        If TypeCheck(Current(DefPos), nbFire) Then
            EffectTook = 0
            If TempMove.Power = 0 Then
                Call AddMessage(45, nbActive, DefPos)
                GoTo EndProcess
            End If
        End If
    'Poison-types can't be poisoned
    Case 81, 115, 119
        If TypeCheck(Current(DefPos), nbPoison) Then
            EffectTook = 0
            If TempMove.Power = 0 And TMATCH > 0 Then
                Call AddMessage(45, nbActive, DefPos)
                GoTo EndProcess
            End If
        End If
    'Ghosts can't be Glared at.  I have no freaking idea why not.
    Case 18
        If BattleMode <> nbRBYBattle Then
            If M = 81 And TypeCheck(Current(DefPos), nbGhost) Then
                EffectTook = 0
                Call AddMessage(45, nbActive, DefPos)
                GoTo EndProcess
            End If
        End If
    End Select
    'Steel-types are immune to all Poison attacks
    'Ground-types are immune to all Electric atttacks
    If TMATCH = 0 And Not TempMove.SelfMove And TempMove.Power = 0 Then
        If (TempMove.Type = nbPoison And TypeCheck(Current(DefPos), nbSteel)) _
        Or (TempMove.Type = nbElectr And TypeCheck(Current(DefPos), nbGround)) Then
            Call AddMessage(45, nbActive, DefPos)
            EffectTook = 0
            GoTo EndProcess
        End If
    End If
     
    'ATTACK FAILED
    If Damage(1) = 0 And EffectTook = 0 And (TMATCH > 0 Or TempMove.SelfMove) Then
        If TempMove.SpecialEffect = 47 Then 'OHKO
            Call AddMessage(231, nbActive, DefPos)
        ElseIf TempMove.SpecialEffect = 46 Then 'Leech Seed
            Call AddMessage(210, nbActive, DefPos)
        ElseIf TempMove.SpecialEffect = 93 And BattleMode = nbRBYBattle Then 'RBY Roar/WW
            Call AddMessage(231, nbActive, DefPos)
        Else
            Call AddMessage(38)
        End If
        BCondition(AtkPos).Charging = False
        BCondition(AtkPos).Recharging = False
        BCondition(AtkPos).FuryCutter = 0
        BCondition(AtkPos).Rollout = 0
        BCondition(AtkPos).SemiInvul = 0
        BCondition(AtkPos).ProtectPercent = 0
        GoTo EndProcess
    End If
    
    'ATTACK MISSED
    If Hit = False Then
        If MultiTarget Then
            Call AddMessage(210, nbActive, DefPos)
        Else
            Call AddMessage(37, nbActive, AtkPos)
        End If
        BCondition(AtkPos).Charging = False
        BCondition(AtkPos).FuryCutter = 0
        BCondition(AtkPos).Recharging = False
        BCondition(AtkPos).Rollout = 0
        BCondition(AtkPos).SemiInvul = 0
        If BattleMode <> nbAdvBattle Then BCondition(AtkPos).RageCounter = 0
        If BCondition(AtkPos).StuckMove > 0 Then
            If Moves(BCondition(AtkPos).StuckMove).SpecialEffect = 121 Then
                BCondition(AtkPos).StuckMove = 0
            End If
        End If
        If TempMove.SpecialEffect = 61 Then
            Call AddMessage(221, nbActive, AtkPos)
            X = Damage(1) \ 8
            If X = 0 Or BattleMode = nbRBYBattle Then X = 1
            Current(AtkPos).HP = Current(AtkPos).HP - X
            If BattleMode = nbRBYBattle Then
                BCondition(DefPos).LastPhDamage = X
                BCondition(DefPos).LastPhHitter = AtkPos
            End If
        End If
        GoTo EndProcess
    End If
    
    'WATER/VOLT ABSORB
    If TempMove.Power > 0 Then
        If (Current(DefPos).Attribute = nbWaterAbsorb And TempMove.Type = nbWater) _
        Or (Current(DefPos).Attribute = nbVoltAbsorb And TempMove.Type = nbElectr) Then
            If Current(DefPos).HP = Current(DefPos).MaxHP Then
                Call AddMessage(140, nbActive, DefPos, nbTrait, Current(DefPos).Attribute, nbMove, M)
            Else
                Call AddMessage(136, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
                Current(DefPos).HP = Cap(Current(DefPos).HP + Current(DefPos).MaxHP \ 4, Current(DefPos).MaxHP)
            End If
            EffectTook = 0
            BCondition(AtkPos).Recharging = False
            BCondition(AtkPos).SemiInvul = 0
            If BCondition(AtkPos).StuckMove > 0 Then
                If Moves(BCondition(AtkPos).StuckMove).SpecialEffect = 121 Then
                    BCondition(AtkPos).StuckMove = 0
                End If
            End If
            GoTo EndProcess
        End If
    End If
    
    'TYPE IMMUNITY (DAMAGING)
    If TMATCH = 0 And Not TempMove.SelfMove And TempMove.Power > 0 Then
        'RBY Wrap/Bind stupidity
        If BattleMode = nbRBYBattle And TempMove.SpecialEffect = 15 And Hit Then
            If BCondition(AtkPos).RepeatCount = 0 Then
                Call AddMessage(45, nbActive, DefPos)
            End If
            TempMove.Power = 0
        Else
            Call AddMessage(45, nbActive, DefPos)
            EffectTook = 0
            GoTo EndProcess
        End If
    End If
    
    'FLASH FIRE
    If Current(DefPos).Attribute = nbFlashFire And TempMove.Type = nbFire And Not TempMove.SelfMove Then
        If BCondition(DefPos).FlashFire Then
            Call AddMessage(158, nbActive, DefPos, nbTrait, Current(DefPos).Attribute, nbMove, M)
        Else
            Call AddMessage(150, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
        End If
        BCondition(DefPos).FlashFire = True
        EffectTook = 0
        GoTo EndProcess
    End If
    
    'SAFEGUARD
    If TeamBCondition(DefTeam).SafeguardCount > 0 Then
        Select Case TempMove.SpecialEffect
        Case 17, 18, 125, 22, 39, 124, 64, 81, 115, 119, 156
            EffectTook = 0
            If TempMove.Power = 0 Then
                Call AddMessage(79, nbActive, DefPos, nbMove, 173)
                GoTo EndProcess
            End If
        End Select
    End If
    
    'Present freakiness
    If TempMove.SpecialEffect = 82 And EffectTook = 2 Then TempMove.Power = 0

    'If it's gotten this far, the attack goes through.
    '-----ATTACK EXECUTION-----

    'Call DoAnim(BattleWindow, nbaMove + M, AtkPos, DefPos)

    For Z = 1 To HitCount
        If TempMove.Power > 0 Then
            'Inflict damage
            B = 0
            If BCondition(DefPos).Substitute <= 0 Then
                HitSubstitute = False
                Damaged = True
                Current(DefPos).HP = Current(DefPos).HP - Damage(Z)
                If Current(DefPos).HP <= 0 Then
                    MultihitExit = True
                    Damage(Z) = Damage(Z) + Current(DefPos).HP
                    If BCondition(DefPos).Enduring Then
                        B = B + 1
                    ElseIf TempMove.SpecialEffect = 45 Then
                        B = B + 2
                    ElseIf FocusBand Then
                        B = B + 4
                    End If
                    If B > 0 Then
                        Current(DefPos).HP = 1
                        Damage(Z) = Damage(Z) - 1
                    End If
                End If
                MultihitDamage = MultihitDamage + Damage(Z)
                If TempMove.SpecialPercent > 0 And Current(DefPos).HP <= 0 Then EffectTook = 0
                If BMessStyle = 1 And Not Multihitter Then B = B + 8
                If BCondition(DefPos).BideCount > 0 Then BCondition(DefPos).BideDamage = BCondition(DefPos).BideDamage + Damage(Z)
            Else
                HitSubstitute = True
                BCondition(DefPos).Substitute = BCondition(DefPos).Substitute - Damage(Z)
                B = B + 1
                If BCondition(DefPos).Substitute <= 0 Then
                    Damage(Z) = Damage(Z) + BCondition(DefPos).Substitute
                    BCondition(DefPos).Substitute = 0
                    B = B + 2
                    If BattleMode = nbRBYBattle Then MultihitExit = True
                End If
                If BCondition(DefPos).BideCount > 0 And BattleMode = nbRBYBattle Then BCondition(DefPos).BideDamage = BCondition(DefPos).BideDamage + Damage(Z)
            End If
            
            'Multihitter's Hit message
            If Multihitter Then
                If TempMove.SpecialEffect = 12 Then
                    A = Current(AtkPos).TeamNumber
                    TempString = Current(AtkPos).Nickname
                    TempString2 = Current(AtkPos).Name
                    X = 0
                    For Y = 1 To 6
                        If PKMN(AtkTeam, Y).Condition = 1 Then
                            X = X + 1
                        End If
                        If Z = X Then Exit For
                    Next Y
                    'This freakiness is because AddMessage can only read the current poke,
                    'and letting it read the whole team would be nasty to code.
                    Current(AtkPos).TeamNumber = Y
                    Current(AtkPos).Nickname = PKMN(AtkTeam, Y).Nickname
                    Current(AtkPos).Name = PKMN(AtkTeam, Y).Name
                    Call AddMessage(294, nbActive, AtkPos)
                    Current(AtkPos).TeamNumber = A
                    Current(AtkPos).Nickname = TempString
                    Current(AtkPos).Name = TempString2
                ElseIf Not CHitArray(Z) Then
                    Call AddMessage(9)
                End If
            End If
            
            'A few more messages
            If HitSubstitute Then
                If (B And 1) = 1 Then Call AddMessage(100, nbActive, DefPos)
                If (B And 2) = 2 Then Call AddMessage(101, nbActive, DefPos)
            Else
                If (B And 8) = 8 Then Call DoDamageMessage(DefPos, Damage(Z))
                If (B And 1) = 1 Then Call AddMessage(94, nbActive, DefPos)
                If (B And 4) = 4 Then Call AddMessage(132, nbActive, DefPos, nbItem, Current(DefPos).Item)
            End If
        
            'Critical Hit message
            If CHitArray(Z) Then Call AddMessage(5)
            
            'RBY's OHKO message
            If BattleMode = nbRBYBattle And TempMove.SpecialEffect = 47 Then Call AddMessage(6)
            
            'Type Effectiveness message
            If Not SkipTypeMessage And Not Multihitter Then
                If TMATCH < 1 Then
                    Call AddMessage(7)
                ElseIf TMATCH > 1 Then
                    Call AddMessage(8)
                End If
            End If
            
            'Unfreeze if hit with Fire
            If TempMove.Type = 2 And Current(DefPos).Condition = 7 And TempMove.SpecialEffect <> 62 Then
                Current(DefPos).Condition = 1
                Call AddMessage(185, nbActive, DefPos)
            End If
            
            'Set LastXXXX variables
            If Not HitSubstitute Then
                BCondition(DefPos).Damaged = True 'For Focus Punch and Revenge
                BCondition(DefPos).LastAnyHitter = AtkPos 'For Bide
            End If
            'Set last damages for Counter/Mirror Coat
            If Not HitSubstitute Then
                X = Damage(Z)
                If AtkTeam = DefTeam Then X = 0
                If BattleMode = nbRBYBattle And (TempMove.Type = 7 Or TempMove.Type = 1) Then
                    BCondition(DefPos).LastPhDamage = X
                    BCondition(DefPos).LastPhHitter = AtkPos
                ElseIf BattleMode <> nbRBYBattle And AttackBase(Moves(M).Type) = nbPysicalBased Then
                    BCondition(DefPos).LastPhDamage = X
                    BCondition(DefPos).LastPhHitter = AtkPos
                Else
                    BCondition(DefPos).LastSpDamage = X
                    BCondition(DefPos).LastSpHitter = AtkPos
                End If
            End If
        End If
        BCondition(DefPos).MirrorPos = AtkPos 'For Mirror Move
        BCondition(DefPos).LastAttack = M 'For Conversion 2
        If BCondition(AtkPos).LastMoveSlot > 0 Then
            If Current(AtkPos).Move(BCondition(AtkPos).LastMoveSlot) = M Then
                BCondition(DefPos).MirrorMove = M 'For Mirror Move
            End If
        End If
        Select Case TempMove.SpecialEffect
        Case 10, 19, 119: If BattleMode = nbAdvBattle Then BCondition(DefPos).DBPos = 0 'For the DBond glitch
        Case 84: BCondition(DefPos).MirrorMove = 0 'For the Psych Up/Mirror Move glitch
        Case Else: If TempMove.Power > 0 Then BCondition(DefPos).DBPos = AtkPos  'For Destiny Bond and Grudge
        End Select
        
        If Not Multihitter Then
            'For Mist/Clear Body/White Smoke/Hyper Cutter/Keen Eye
            Select Case TempMove.SpecialEffect
            Case 2, 9, 20, 21, 25, 27, 50, 96, 111, 138, 150, 151
                If (BCondition(DefPos).Mist Or TeamBCondition(DefTeam).MistCount > 0) And EffectTook > 0 Then
                    EffectTook = 0
                    If TempMove.Power = 0 Then Call AddMessage(96, nbActive, DefPos)
                ElseIf Current(DefPos).Attribute = nbClearBody Or Current(DefPos).Attribute = nbWhiteSmoke Then
                    EffectTook = 0
                    If TempMove.Power = 0 Then Call AddMessage(148, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
                ElseIf Current(DefPos).Attribute = nbHyperCutter And (TempMove.SpecialEffect = 9 Or TempMove.SpecialEffect = 21) Then
                    EffectTook = 0
                    If TempMove.Power = 0 Then Call AddMessage(149, nbActive, DefPos, nbTrait, Current(DefPos).Attribute, nbStat, 2)
                ElseIf Current(DefPos).Attribute = nbKeenEye And TempMove.SpecialEffect = 50 Then
                    EffectTook = 0
                    If TempMove.Power = 0 Then Call AddMessage(149, nbActive, DefPos, nbTrait, Current(DefPos).Attribute, nbStat, 8)
                End If
            End Select
            
            'MOVE EFFECTS
            If EffectTook > 0 Then
                Select Case TempMove.SpecialEffect
                'Gain health = 1/2 damage
                Case 1
                    X = Damage(1) \ 2
                    If X = 0 Then X = 1
                    If Current(DefPos).Attribute = nbLiquidOoze Then
                        Call AddMessage(160, nbActive, AtkPos, nbTrait, Current(DefPos).Attribute)
                        Current(AtkPos).HP = Current(AtkPos).HP - X
                    Else
                        Call AddMessage(215, nbActive, DefPos)
                        Current(AtkPos).HP = Current(AtkPos).HP + X
                    End If
                'Defense - 1
                Case 2
                    Call ModifyPKMN(DefPos, 2, -1)
                'Defense + 2
                Case 3
                    Call ModifyPKMN(AtkPos, 2, 2)
                'Speed + 2
                Case 5
                    Call ModifyPKMN(AtkPos, 7, 2)
                'SDefense + 2
                Case 6
                    Call ModifyPKMN(AtkPos, 4, 2)
                'All + 1
                Case 7
                    Call ModifyPKMN(AtkPos, 1, 1)
                    Call ModifyPKMN(AtkPos, 2, 1)
                    Call ModifyPKMN(AtkPos, 7, 1)
                    Call ModifyPKMN(AtkPos, 3, 1)
                    Call ModifyPKMN(AtkPos, 4, 1)
                'Attract
                Case 8
                    BCondition(DefPos).Attract = AtkPos
                    Call AddMessage(196, nbActive, DefPos)
                'Attack - 1
                Case 9
                    Call ModifyPKMN(DefPos, 1, -1)
                'Baton Pass
                Case 11
                    BCondition(AtkPos).BatonPassing = True
                'Belly Drum
                Case 13
                    If Current(AtkPos).HP <= Current(AtkPos).MaxHP \ 2 Then
                        If BattleMode = nbGSCBattle And Not StadiumMode Then Call ModifyPKMN(AtkPos, 1, 2, True)
                        Call AddMessage(38)
                    ElseIf BCondition(AtkPos).AttackChange = 6 Then
                        Call AddMessage(38)
                    Else
                        Call AddMessage(249, nbActive, AtkPos)
                        Call ModifyPKMN(AtkPos, 1, 12, True)
                        Call ChangeHP(AtkPos, -2)
                    End If
                'Wrap
                Case 15
                    If BattleMode = nbRBYBattle Then
                        If BCondition(AtkPos).RepeatCount > 0 Then
                            BCondition(AtkPos).RepeatCount = BCondition(AtkPos).RepeatCount - 1
                            If BCondition(AtkPos).RepeatCount = 0 Then BCondition(AtkPos).RepeatMove = 0
                        Else
                            BCondition(AtkPos).RepeatMove = M
                            BCondition(AtkPos).RepeatCount = EffectTook
                        End If
                    Else
                        If BCondition(DefPos).RepeatMove = 0 Then
                            BCondition(DefPos).RepeatMove = M
                            BCondition(DefPos).RepeatPos = AtkPos
                            BCondition(DefPos).RepeatCount = EffectTook
                            Select Case M
                            Case 16 'Bind
                                Call AddMessage(202, nbActive, DefPos, nbActive, AtkPos, nbMove, M)
                            Case 26 'Clamp
                                Call AddMessage(206, nbActive, AtkPos, nbActive, DefPos)
                            Case 66, 246 'Fire Spin and Whirlpool
                                Call AddMessage(203, nbActive, DefPos)
                            Case 250 'Wrap
                                Call AddMessage(205, nbActive, DefPos, nbActive, AtkPos)
                            Case 324 'Sand Tomb
                                Call AddMessage(204, nbActive, DefPos, nbMove, M)
                            End Select
                        End If
                    End If
                'Flinch
                Case 16, 101, 137
                    BCondition(DefPos).Flinching = True
                'Freeze
                Case 17
                    If Current(DefPos).Condition = 1 _
                    And CurrentWeather <> nbSunny _
                    And Current(DefPos).Attribute <> nbMagmaArmor _
                    And Not RuleCheck(DefTeam, 3) Then
                        Current(DefPos).Condition = 7
                        Current(DefPos).ConditionCount = 0
                        Call AddMessage(183, nbActive, DefPos)
                        Call ClearStuckMove(DefPos)
                    End If
                'Paralyze
                Case 18, 125
                    Current(DefPos).Condition = 6
                    Current(DefPos).ConditionCount = 0
                    BCondition(DefPos).ParalyzePenalty = True
                    Call AddMessage(187, nbActive, DefPos)
                    Call DoSynchronize(AtkPos, DefPos)
                'Speed - 1
                Case 20
                    Call ModifyPKMN(DefPos, 7, -1)
                'Attack - 2
                Case 21
                    Call ModifyPKMN(DefPos, 1, -2)
                'Confusion
                Case 22
                    BCondition(DefPos).Confuse = True
                    BCondition(DefPos).ConfuseCounter = EffectTook
                    Call AddMessage(193, nbActive, DefPos)
                'Conversion and Conversion2
                Case 23, 24
                    If BattleMode = nbRBYBattle Then
                        Current(AtkPos).Type1 = EffectParam Mod 16
                        Current(AtkPos).Type2 = EffectParam \ 16
                    Else
                        Current(AtkPos).Type2 = nbNoType
                        Current(AtkPos).Type1 = EffectParam
                        Call AddMessage(220, nbActive, AtkPos, nbType, EffectParam)
                    End If
                'Speed - 2
                Case 25
                    Call ModifyPKMN(DefPos, 7, -2)
                'SDefense - 1
                Case 27
                    Call ModifyPKMN(DefPos, 4, -1)
                'Curse
                Case 28
                    If TypeCheck(Current(AtkPos), nbGhost) Then
                        Call ChangeHP(AtkPos, -2)
                        BCondition(DefPos).Curse = True
                        Call AddMessage(200, nbActive, AtkPos, nbActive, DefPos)
                    Else
                        Call ModifyPKMN(AtkPos, 1, 1)
                        Call ModifyPKMN(AtkPos, 2, 1)
                        Call ModifyPKMN(AtkPos, 7, -1)
                    End If
                'Defense + 1
                Case 29
                    Call ModifyPKMN(AtkPos, 2, 1)
                    If M = 40 Then BCondition(AtkPos).DefenseCurl = True
                'Destiny Bond
                Case 30
                    Call AddMessage(240, nbActive, AtkPos)
                    BCondition(AtkPos).DestinyBond = True
                'Protect/Detect
                Case 31
                    Call AddMessage(41, nbActive, AtkPos)
                    BCondition(AtkPos).Protected = True
                'Disable
                Case 33
                    BCondition(DefPos).DisableCount = EffectTook
                    BCondition(DefPos).DisabledMove = EffectParam
                    Call AddMessage(234, nbActive, DefPos, nbMove, Current(DefPos).Move(EffectParam))
                'Evasion + 1
                Case 34
                    Call ModifyPKMN(AtkPos, 6, 1)
                    If M = 126 Then BCondition(AtkPos).Minimize = True
                'Receive 1/4 damage
                Case 35
                    If Current(AtkPos).Attribute <> nbRockHead Or M = 210 Then
                        X = Damage(1) \ IIf(M = 210, 2, 4)
                        If X = 0 Or HitSubstitute Then X = 1
                        Current(AtkPos).HP = Current(AtkPos).HP - X
                        Call AddMessage(11, nbActive, AtkPos)
                    End If
                'Dream Eater
                Case 37
                    X = Damage(1) \ 2
                    If X = 0 Then X = 1
                    Current(AtkPos).HP = Current(AtkPos).HP + X
                    Call AddMessage(216, nbActive, DefPos)
                'Burn
                Case 39, 124
                    Current(DefPos).Condition = 5
                    Current(DefPos).ConditionCount = EffectTook
                    BCondition(DefPos).BurnPenalty = True
                    Call AddMessage(179, nbActive, DefPos)
                    Call DoSynchronize(AtkPos, DefPos)
                'Encore
                Case 40
                    Call AddMessage(236, nbActive, DefPos)
                    BCondition(DefPos).Encore = True
                    BCondition(DefPos).EncoreDuration = EffectTook
                    BCondition(DefPos).EncoreMove = EffectParam
                'Endure
                Case 41
                    Call AddMessage(93, nbActive, AtkPos)
                    BCondition(AtkPos).Enduring = True
                'Explosion/SelfDestruct
                Case 42
                    If HitSubstitute And BCondition(DefPos).Substitute <= 0 And BattleMode = nbRBYBattle And Not StadiumMode Then
                        Current(AtkPos).HP = BCondition(AtkPos).ExplosionHP
                        BCondition(AtkPos).FaintLast = False
                        SelfLoss(AtkTeam) = False
                    End If
                'Leech Seed
                Case 46
                    If BCondition(DefPos).LeechSeed = 0 Then
                        BCondition(DefPos).LeechSeed = AtkPos
                        Call AddMessage(209, nbActive, DefPos)
                    Else
                        Call AddMessage(210, nbActive, DefPos)
                    End If
                'OHKO
                Case 47
                    If Current(DefPos).HP <= 0 And BattleMode <> nbRBYBattle Then Call AddMessage(6)
                'Accuracy - 1
                Case 50
                    Call ModifyPKMN(DefPos, 5, -1)
                'Dig, Fly, and Dive
                Case 32, 51, 132
                    BCondition(AtkPos).SemiInvul = 0
                'Focus Energy
                Case 52
                    BCondition(AtkPos).FocusEnergy = True
                    Call AddMessage(212, nbActive, AtkPos)
                'Foresight
                Case 53
                    'BCondition(DefPos).EvadeChange = 0
                    BCondition(DefPos).Foresight = True
                    Call AddMessage(247, nbActive, AtkPos, nbActive, DefPos)
                'Future Sight
                Case 56
                    With BCondition(DefPos).FutureSight
                        .count = 3
                        If AttackBase(Moves(M).Type) = nbPysicalBased Then
                            .AttackPower = GetAtk(AtkPos)
                        Else
                            .AttackPower = GetSAtk(AtkPos)
                        End If
                        .Level = Current(AtkPos).Level
                        .AccuracyMod = BCondition(AtkPos).AccuracyChange
                        .CHit = CHit
                        .FSMove = IIf(M = 79, 0, 1)
                        .RandomNumber = EffectParam
                    End With
                    Call AddMessage(255, nbActive, AtkPos)
                'Light Screen
                Case 57
                    TeamBCondition(AtkTeam).LightScreenCount = 5
                    X = 82
                    If ActNum = 4 Then
                        For Y = AtkTeam To 4 Step 2
                            If Current(Y).HP <= 0 Then Exit For
                        Next Y
                        If Y > 4 Then X = 83
                    End If
                    Call AddMessage(X, nbPlayer, AtkTeam, nbMove, M, nbStat, 7)
                'SAttack + 1
                Case 58
                    Call ModifyPKMN(AtkPos, 3, 1)
                'Haze
                Case 59
                    For X = 1 To ActNum
                        BCondition(X).AccuracyChange = 0
                        BCondition(X).AttackChange = 0
                        BCondition(X).DefenseChange = 0
                        BCondition(X).EvadeChange = 0
                        BCondition(X).SAttackChange = 0
                        BCondition(X).SDefenseChange = 0
                        BCondition(X).SpeedChange = 0
                        If BattleMode = nbRBYBattle Then
                            BCondition(X).ParalyzePenalty = False
                            BCondition(X).BurnPenalty = False
                            BCondition(X).LeechSeed = 0
                            BCondition(X).Mist = False
                            BCondition(X).Confuse = False
                            BCondition(X).ConfuseCounter = 0
                            BCondition(X).ToxicCount = 0
                            TeamBCondition(X).ReflectCount = 0
                            TeamBCondition(X).LightScreenCount = 0
                            If X = DefPos Then Current(X).Condition = 1
                        End If
                    Next X
                    Call AddMessage(92)
                'Heal Bell/Aromatherepy
                Case 60
                    If M = 89 Then
                        Call AddMessage(91)
                    Else
                        Call AddMessage(90)
                    End If
                    For X = 1 To NumPoke
                        If M <> 89 Or (PKMN(AtkTeam, X).Attribute <> nbSoundproof And PKMN(AtkTeam, X).Attribute <> nbCacophony) Then
                            If PKMN(AtkTeam, X).Condition < 8 Then PKMN(AtkTeam, X).Condition = 1
                            PKMN(AtkTeam, X).Resting = False
                        End If
                    Next X
                    For X = AtkTeam To ActNum Step 2
                        With Current(X)
                            If M <> 89 Or (.Attribute <> nbSoundproof And .Attribute <> nbCacophony) Then
                                If .Condition < 8 Then
                                    .Condition = 1
                                    .Resting = False
                                    BCondition(AtkPos).BurnPenalty = False
                                    BCondition(AtkPos).ParalyzePenalty = False
                                End If
                            Else
                                Call AddMessage(155, nbActive, X, nbTrait, Current(X).Attribute, nbMove, M)
                            End If
                        End With
                    Next X
                'Hyper Beam
                Case 63
                    BCondition(AtkPos).Recharging = True
                    If HitSubstitute And BCondition(DefPos).Substitute <= 0 And BattleMode = nbRBYBattle Then BCondition(AtkPos).Recharging = False
                'Sleep
                Case 64
                    Current(DefPos).Condition = 4
                    Current(DefPos).ConditionCount = EffectTook
                    Current(DefPos).Resting = False
                    BCondition(DefPos).Nightmare = False
                    Call AddMessage(169, nbActive, DefPos)
                    Call ClearStuckMove(DefPos)
                'Lock-On
                Case 65
                    Call AddMessage(238, nbActive, AtkPos, nbActive, DefPos)
                    BCondition(AtkPos).LockOn = DefPos
                    BCondition(AtkPos).LockOnCount = 2
                'Mean Look
                Case 67
                    BCondition(DefPos).TrapNum = AtkPos
                    Call AddMessage(244, nbActive, DefPos)
                'Attack + 1
                Case 68
                    Call ModifyPKMN(AtkPos, 1, 1)
                'Mimic
                Case 71
                    BCondition(AtkPos).MimicedMove = 0
                    For X = 1 To 4
                        If Current(AtkPos).Move(X) = M Then
                            Current(AtkPos).Move(X) = EffectParam
                            Current(AtkPos).MaxPP(X) = Moves(EffectParam).PP * IIf(BattleMode <> nbRBYBattle Or Not EnabledRule(nbUsePPUps), 1, 1.6)
                            If Not BattleMode = nbRBYBattle Then Current(AtkPos).PP(X) = 5
                            BCondition(AtkPos).MimicedMove = EffectParam
                            Call AddMessage(228, nbActive, AtkPos, nbMove, EffectParam)
                        End If
                    Next X
                'Mist
                Case 74
                    If BattleMode = nbAdvBattle Then
                        TeamBCondition(AtkTeam).MistCount = 5
                        Call AddMessage(95, nbPlayer, AtkTeam)
                    Else
                        BCondition(AtkPos).Mist = True
                        Call AddMessage(296, nbActive, AtkPos)
                    End If
                'Synthesis/Moonlight/Morning Sun
                Case 75
                    If Current(AtkPos).HP = Current(AtkPos).MaxHP Then
                        Call AddMessage(81, nbActive, AtkPos)
                    Else
                        If CurrentWeather = nbNoWeather Then
                            X = 2
                        ElseIf CurrentWeather = nbSunny Then
                            X = 1
                        Else
                            X = 4
                        End If
                        Call ChangeHP(AtkPos, X)
                        Call AddMessage(80, nbActive, AtkPos)
                    End If
                'Nightmare
                Case 77
                    BCondition(DefPos).Nightmare = True
                    Call AddMessage(198, nbActive, DefPos)
                'Pain Split
                Case 78
                    Temp = (Current(DefPos).HP + Current(AtkPos).HP) \ 2
                    Current(DefPos).HP = Temp
                    Current(AtkPos).HP = Temp
                    Call AddMessage(291)
                'Outrage/Thrash/Petal Dance
                Case 79
                    If Current(AtkPos).Condition <> nbSlp And BattleMode <> nbRBYBattle Then
                        If BCondition(AtkPos).StuckMove = 0 Then
                            BCondition(AtkPos).StuckMove = M
                            BCondition(AtkPos).StuckCount = EffectTook
                            BCondition(AtkPos).OutrageConfuse = EffectParam
                        End If
                        If BattleMode <> nbAdvBattle Then
                            BCondition(AtkPos).StuckCount = BCondition(AtkPos).StuckCount - 1
                            If BCondition(AtkPos).StuckCount = 0 Then
                                BCondition(AtkPos).StuckMove = 0
                                If BCondition(AtkPos).Confuse = False And TeamBCondition(AtkPos).SafeguardCount = 0 Then
                                    BCondition(AtkPos).Confuse = True
                                    BCondition(AtkPos).ConfuseCounter = BCondition(AtkPos).OutrageConfuse
                                    Call AddMessage(195, nbActive, AtkPos)
                                End If
                            End If
                        End If
                    End If
                'Pay Day
                Case 80
                    Call AddMessage(290)
                'Poison
                Case 81, 119, 156
                    Current(DefPos).Condition = 2
                    Call AddMessage(174, nbActive, DefPos)
                    Call DoSynchronize(AtkPos, DefPos)
                'Present
                Case 82
                    If EffectTook = 2 Then
                        If Current(DefPos).HP = Current(DefPos).MaxHP Then
                            Call AddMessage(81, nbActive, DefPos)
                        Else
                            Call AddMessage(214, nbActive, DefPos)
                            Call ChangeHP(DefPos, 4)
                        End If
                    End If
                'Perish Song
                Case 83
                    If BattleMode = nbGSCBattle Then
                        If BCondition(DefPos).PerishSong = 0 And BCondition(AtkPos).PerishSong = 0 Then
                            BCondition(DefPos).PerishSong = 4
                            BCondition(AtkPos).PerishSong = 4
                            Call AddMessage(292)
                        Else
                            Call AddMessage(38)
                        End If
                    Else
                        Y = 0
                        For X = 1 To ActNum
                            If BCondition(X).PerishSong = 0 And Current(X).HP > 0 Then Y = Y + 1
                        Next X
                        If Y = 0 Then
                            Call AddMessage(38)
                        Else
                            Call AddMessage(292)
                        End If
                        For X = 1 To ActNum
                            If BCondition(X).PerishSong = 0 And Current(X).HP > 0 Then
                                If Current(X).Attribute = nbSoundproof Or Current(X).Attribute = nbCacophony Then
                                    Call AddMessage(155, nbActive, X, nbTrait, Current(X).Attribute, nbMove, M)
                                Else
                                    BCondition(X).PerishSong = 4
                                End If
                            End If
                        Next X
                    End If
                'Psych Up
                Case 84
                    BCondition(AtkPos).AttackChange = BCondition(DefPos).AttackChange
                    BCondition(AtkPos).DefenseChange = BCondition(DefPos).DefenseChange
                    BCondition(AtkPos).SpeedChange = BCondition(DefPos).SpeedChange
                    BCondition(AtkPos).SAttackChange = BCondition(DefPos).SAttackChange
                    BCondition(AtkPos).SDefenseChange = BCondition(DefPos).SDefenseChange
                    BCondition(AtkPos).EvadeChange = BCondition(DefPos).EvadeChange
                    BCondition(AtkPos).AccuracyChange = BCondition(DefPos).AccuracyChange
                    Call AddMessage(250, nbActive, AtkPos, nbActive, DefPos)
                'Rage
                Case 87
                    If BattleMode = nbRBYBattle Then BCondition(AtkPos).StuckMove = 156
                    If BCondition(AtkPos).RageCounter = 0 Then BCondition(AtkPos).RageCounter = 1
                'Rain Dance
                Case 88
                    RealWeather = nbRaining
                    WeatherCount = 5
                    Call AddMessage(103)
                'Rapid Spin
                Case 89
                    If BCondition(DefPos).Substitute <= 0 Then
                        If BCondition(AtkPos).RepeatMove > 0 Then
                            Call AddMessage(251, nbActive, AtkPos, nbActive, BCondition(AtkPos).RepeatPos, nbMove, BCondition(AtkPos).RepeatMove)
                            BCondition(AtkPos).RepeatMove = 0
                        End If
                        If BCondition(AtkPos).LeechSeed <> 0 Then
                            BCondition(AtkPos).LeechSeed = 0
                            Call AddMessage(252, nbActive, AtkPos)
                        End If
                        If TeamBCondition(AtkTeam).Spikes > 0 Then
                            TeamBCondition(AtkTeam).Spikes = 0
                            Call AddMessage(253, nbActive, AtkPos)
                        End If
                    End If
                'Recover
                Case 90
                    With Current(AtkPos)
                        If .HP = .MaxHP Then
                            If BattleMode = nbRBYBattle Then
                                Call AddMessage(38)
                            Else
                                Call AddMessage(81, nbActive, AtkPos)
                            End If
                        ElseIf BattleMode = nbRBYBattle And (.MaxHP - .HP) Mod 256 = 255 Then
                            Call AddMessage(38)
                        Else
                            Call AddMessage(80, nbActive, AtkPos)
                            Call ChangeHP(AtkPos, 2)
                        End If
                    End With
                'Reflect
                Case 91
                    TeamBCondition(AtkTeam).ReflectCount = 5
                    X = 82
                    If ActNum = 4 Then
                        For Y = AtkTeam To 4 Step 2
                            If Current(Y).HP <= 0 Then Exit For
                        Next Y
                        If Y > 4 Then X = 83
                    End If
                    Call AddMessage(X, nbPlayer, AtkTeam, nbMove, M, nbStat, 3)
                'Rest
                Case 92
                    With Current(AtkPos)
                        If UproarCheck(AtkPos) <> 0 Then
                            Call AddMessage(229, nbActive, AtkPos)
                        ElseIf .Attribute = nbInsomnia Or .Attribute = nbVitalSpirit Then
                            Call AddMessage(230, nbActive, AtkPos, nbTrait, .Attribute)
                        ElseIf .Condition = 4 And BattleMode = nbAdvBattle Then
                            Call AddMessage(171, nbActive, AtkPos)
                        ElseIf .HP = .MaxHP Then
                            If BattleMode = nbRBYBattle Then
                                Call AddMessage(38)
                            Else
                                Call AddMessage(81, nbActive, AtkPos)
                            End If
                        ElseIf BattleMode = nbRBYBattle And (.MaxHP - .HP) Mod 256 = 255 Then
                            Call AddMessage(38)
                        Else
                            If .Condition <> 4 Then BCondition(AtkPos).Nightmare = False
                            .Condition = 4
                            .ConditionCount = IIf(BattleMode = nbRBYBattle, 2, 3)
                            .Resting = True
                            Call AddMessage(84, nbActive, AtkPos)
                            Call ClearStuckMove(AtkPos)
                            .HP = .MaxHP
                            If Not BattleMode = nbRBYBattle Then
                                BCondition(AtkPos).ParalyzePenalty = False
                                BCondition(AtkPos).BurnPenalty = False
                            End If
                        End If
                    End With
                'Roar/Whirlwind
                Case 93
                    Call AddMessage(IIf(M = 247, 219, 254), nbActive, DefPos)
                    BCondition(DefPos).TraceNumber = EffectParam
                    X = Current(DefPos).TeamNumber
                    With PKMN(DefTeam, EffectTook)
                        Current(DefPos).Nickname = .Nickname
                        Current(DefPos).TeamNumber = .TeamNumber
                    End With
                    Call AddMessage(285, nbActive, DefPos, nbNumber, Current(DefPos).Level, nbPoke, PKMN(DefTeam, EffectTook).No)
                    Current(DefPos).TeamNumber = X
                    Call DoSwitch(DefPos, EffectTook, , True, True)
                'Safeguard
                Case 94
                    TeamBCondition(AtkTeam).SafeguardCount = 5
                    Call AddMessage(78, nbPlayer, AtkTeam)
                'Sandstorm
                Case 95
                    RealWeather = nbSandstorm
                    WeatherCount = 5
                    Call AddMessage(108)
                'Defense - 2
                Case 96
                    Call ModifyPKMN(DefPos, 2, -2)
                'Sketch
                Case 97
                    For X = 1 To 4
                        If Current(AtkPos).Move(X) = 184 Then
                            Current(AtkPos).Move(X) = EffectParam
                            Current(AtkPos).PP(X) = Moves(EffectParam).PP
                            Current(AtkPos).MaxPP(X) = Moves(EffectParam).PP
                            PKMN(AtkTeam, Current(AtkPos).TeamNumber).Move(X) = EffectParam
                            PKMN(AtkTeam, Current(AtkPos).TeamNumber).PP(X) = Moves(EffectParam).PP
                            PKMN(AtkTeam, Current(AtkPos).TeamNumber).MaxPP(X) = Moves(EffectParam).PP
                            Call AddMessage(239, nbActive, AtkPos, nbMove, EffectParam)
                        End If
                    Next X
                'Spikes
                Case 104
                    TeamBCondition(OtherTeam(AtkTeam)).Spikes = TeamBCondition(OtherTeam(AtkTeam)).Spikes + 1
                    Call AddMessage(245, nbPlayer, OtherTeam(AtkTeam))
                'Spite
                Case 105
                    Call DoPPChange(DefPos, EffectParam, -1 * EffectTook)
                    Call AddMessage(242, nbActive, DefPos, nbMove, Current(DefPos).Move(EffectParam), nbNumber, EffectTook)
                'Splash
                Case 106
                    Call AddMessage(39)
                'Substitute
                Case 107
                    X = Current(AtkPos).MaxHP \ 4
                    If X = 0 Then X = 1
                    If BCondition(AtkPos).Substitute > 0 Then
                        Call AddMessage(99, nbActive, AtkPos)
                    ElseIf Current(AtkPos).HP <= X Then
                        Call AddMessage(102)
                    Else
                        Call AddMessage(98, nbActive, AtkPos)
                        Current(AtkPos).HP = Current(AtkPos).HP - X
                        BCondition(AtkPos).Substitute = Current(AtkPos).MaxHP \ 4
                        BCondition(AtkPos).RepeatMove = 0
                    End If
                'Sunny Day
                Case 108
                    RealWeather = nbSunny
                    WeatherCount = 5
                    Call AddMessage(112)
                'Swagger
                Case 110
                    X = BCondition(DefPos).AttackChange
                    Call ModifyPKMN(DefPos, 1, 2)
                    If (X <> BCondition(DefPos).AttackChange Or BattleMode = nbAdvBattle) _
                    And BCondition(DefPos).Substitute = 0 _
                    And BCondition(DefPos).Confuse = False Then
                        If Current(DefPos).Attribute = nbOwnTempo Then
                            Call AddMessage(145, nbActive, DefPos, nbTrait, nbOwnTempo)
                        ElseIf TeamBCondition(DefTeam).SafeguardCount <> 0 Then
                            Call AddMessage(79, nbActive, DefPos, nbMove, 173)
                        Else
                            BCondition(DefPos).Confuse = True
                            BCondition(DefPos).ConfuseCounter = EffectTook
                            Call AddMessage(193, nbActive, DefPos)
                        End If
                    End If
                'Evade - 1
                Case 111
                    Call ModifyPKMN(DefPos, 6, -1)
                'Attack + 2
                Case 112
                    Call ModifyPKMN(AtkPos, 1, 2)
                'Teleport
                Case 113
                    Call AddMessage(38)
                'Thief
                Case 114
                    If Current(AtkPos).Item = nbNoItem And Current(DefPos).Item > nbNoItem Then
                        If Current(DefPos).Attribute = nbStickyHold Then
                            Call AddMessage(158, nbActive, DefPos, nbTrait, nbStickyHold, nbMove, M)
                        Else
                            Current(AtkPos).Item = Current(DefPos).Item
                            Current(AtkPos).ItemEffect = EffectParam
                            Current(DefPos).Item = nbNoItem
                            BCondition(AtkPos).ChoiceBandSlot = 0
                            BCondition(DefPos).ChoiceBandSlot = 0
                            Call AddMessage(243, nbActive, AtkPos, nbActive, DefPos, nbItem, Current(AtkPos).Item)
                        End If
                    End If
                'Toxic
                Case 115
                    Current(DefPos).Condition = 3
                    BCondition(DefPos).ToxicCount = 0
                    Call AddMessage(178, nbActive, DefPos)
                    Call DoSynchronize(AtkPos, DefPos)
                'Transform
                Case 116
                    X = Current(AtkPos).TeamNumber
                    'Y = Current(AtkPos).Attribute
                    Temp2 = Current(AtkPos).ConditionCount
                    Current(AtkPos) = Current(DefPos)
                    With Current(AtkPos)
                        '.Attribute = Y
                        .TeamNumber = X
                        .Nickname = PKMN(AtkTeam, .TeamNumber).Nickname
                        .Name = PKMN(AtkTeam, .TeamNumber).Name
                        .Item = PKMN(AtkTeam, .TeamNumber).Item
                        .ItemEffect = PKMN(AtkTeam, .TeamNumber).ItemEffect
                        .MaxHP = PKMN(AtkTeam, .TeamNumber).MaxHP
                        .HP = PKMN(AtkTeam, .TeamNumber).HP
                        .Level = PKMN(AtkTeam, .TeamNumber).Level
                        .Gender = PKMN(AtkTeam, .TeamNumber).Gender
                        .DV_HP = PKMN(AtkTeam, .TeamNumber).DV_HP
                        .DV_Atk = PKMN(AtkTeam, .TeamNumber).DV_Atk
                        .DV_Def = PKMN(AtkTeam, .TeamNumber).DV_Def
                        .DV_Spd = PKMN(AtkTeam, .TeamNumber).DV_Spd
                        .DV_SAtk = PKMN(AtkTeam, .TeamNumber).DV_SAtk
                        .DV_SDef = PKMN(AtkTeam, .TeamNumber).DV_SDef
                        .EV_HP = PKMN(AtkTeam, .TeamNumber).EV_HP
                        .EV_Atk = PKMN(AtkTeam, .TeamNumber).EV_Atk
                        .EV_Def = PKMN(AtkTeam, .TeamNumber).EV_Def
                        .EV_Spd = PKMN(AtkTeam, .TeamNumber).EV_Spd
                        .EV_SAtk = PKMN(AtkTeam, .TeamNumber).EV_SAtk
                        .EV_SDef = PKMN(AtkTeam, .TeamNumber).EV_SDef
                        Select Case .GameVersion
                        Case 2, 3
                            .Attack = GetAdvStat(.BaseAttack, .DV_Atk, .EV_Atk, .Level, Nature(.NatureNum).StatChg(1))
                            .Defense = GetAdvStat(.BaseDefense, .DV_Def, .EV_Def, .Level, Nature(.NatureNum).StatChg(2))
                            .Speed = GetAdvStat(.BaseSpeed, .DV_Spd, .EV_Spd, .Level, Nature(.NatureNum).StatChg(3))
                            .SpecialAttack = GetAdvStat(.BaseSAttack, .DV_SAtk, .EV_SAtk, .Level, Nature(.NatureNum).StatChg(4))
                            .SpecialDefense = GetAdvStat(.BaseSDefense, .DV_SDef, .EV_SDef, .Level, Nature(.NatureNum).StatChg(5))
                        Case 0, 5
                            .Attack = GetStat(.Level, .BaseAttack, .DV_Atk)
                            .Defense = GetStat(.Level, .BaseDefense, .DV_Def)
                            .Speed = GetStat(.Level, .BaseSpeed, .DV_Spd)
                            .SpecialAttack = GetStat(.Level, .BaseSpecial, .DV_SAtk)
                            .SpecialDefense = GetStat(.Level, .BaseSpecial, .DV_SAtk)
                        Case Else
                            .Attack = GetStat(.Level, .BaseAttack, .DV_Atk)
                            .Defense = GetStat(.Level, .BaseDefense, .DV_Def)
                            .Speed = GetStat(.Level, .BaseSpeed, .DV_Spd)
                            .SpecialAttack = GetStat(.Level, .BaseSAttack, .DV_SAtk)
                            .SpecialDefense = GetStat(.Level, .BaseSDefense, .DV_SAtk)
                        End Select
                        .Condition = PKMN(AtkTeam, .TeamNumber).Condition
                        .ConditionCount = Temp2
                        .Resting = PKMN(AtkTeam, .TeamNumber).Resting
                        .RecycleItem = PKMN(AtkTeam, .TeamNumber).RecycleItem
                        .Image = ChooseImage(Current(AtkPos), GameVersion(AtkTeam))
                        For X = 1 To 4
                            If .Move(X) > 0 Then .PP(X) = Cap(Moves(.Move(X)).PP, 5)
                        Next X
                    End With
                    BCondition(AtkPos).AttackChange = BCondition(DefPos).AttackChange
                    BCondition(AtkPos).DefenseChange = BCondition(DefPos).DefenseChange
                    BCondition(AtkPos).SpeedChange = BCondition(DefPos).SpeedChange
                    BCondition(AtkPos).SAttackChange = BCondition(DefPos).SAttackChange
                    BCondition(AtkPos).SDefenseChange = BCondition(DefPos).SDefenseChange
                    BCondition(AtkPos).EvadeChange = BCondition(DefPos).EvadeChange
                    BCondition(AtkPos).AccuracyChange = BCondition(DefPos).AccuracyChange
                    If BattleMode = nbRBYBattle Then
                        BCondition(AtkPos).BurnPenalty = BCondition(DefPos).BurnPenalty
                        BCondition(AtkPos).ParalyzePenalty = BCondition(DefPos).ParalyzePenalty
                    End If
                    Call AddMessage(232, nbActive, AtkPos, nbActive, DefPos)
                    BCondition(AtkPos).TransformedTo = Current(DefPos).TeamNumber
                    If AtkTeam = DefTeam Then BCondition(AtkPos).TransformedTo = BCondition(AtkPos).TransformedTo + 6
                'Tri Attack
                Case 117
                    If Current(DefPos).Condition = 1 And TeamBCondition(DefTeam).SafeguardCount = 0 Then
                        Select Case EffectTook
                            Case 1
                                If Current(DefPos).Attribute <> nbLimber Then
                                    Current(DefPos).Condition = 6
                                    Current(DefPos).ConditionCount = 0
                                    BCondition(DefPos).ParalyzePenalty = True
                                    Call AddMessage(187, nbActive, DefPos)
                                    Call DoSynchronize(AtkPos, DefPos)
                                End If
                            Case 2
                                If Current(DefPos).Attribute <> nbMagmaArmor Then
                                    Current(DefPos).Condition = 7
                                    Current(DefPos).ConditionCount = 0
                                    Call AddMessage(183, nbActive, DefPos)
                                    Call ClearStuckMove(DefPos)
                                End If
                            Case 3
                                If Current(DefPos).Attribute <> nbWaterVeil Then
                                    Current(DefPos).Condition = 5
                                    Current(DefPos).ConditionCount = 0
                                    BCondition(DefPos).BurnPenalty = True
                                    Call AddMessage(179, nbActive, DefPos)
                                    Call DoSynchronize(AtkPos, DefPos)
                                End If
                        End Select
                    End If
                    If EffectTook = 4 Then
                        If Current(DefPos).Condition = 7 Then
                            Current(DefPos).Condition = 1
                            Call AddMessage(186, nbActive, DefPos, nbMove, M)
                        End If
                    End If
                'Bulk Up (Atk +1, Def +1)
                Case 127
                    Call ModifyPKMNDouble(AtkPos, 1, 2, 1)
                'Calm Mind (SAtk +1, SDef +1)
                Case 128
                    Call ModifyPKMNDouble(AtkPos, 3, 4, 1)
                'Camouflage
                Case 129
                    Select Case Terrain
                    Case nbShortGrass, nbTallGrass, nbVeryTallGrass
                        X = 5
                    Case nbOcean, nbUnderwater, nbPond
                        X = 3
                    Case nbsAnd
                        X = 9
                    Case nbCave, nbMountain
                        X = 13
                    Case Else
                        X = 1
                    End Select
                    If X = Current(AtkPos).Type1 Then
                        Call AddMessage(38)
                    Else
                        Call AddMessage(220, nbActive, AtkPos, nbType, X)
                        Current(AtkPos).Type1 = X
                        Current(AtkPos).Type2 = nbNoType
                    End If
                'Charge
                Case 130
                    Call AddMessage(259, nbActive, AtkPos)
                    BCondition(AtkPos).ChargeCount = 2
                'Cosmic Power (Def +1, SDef +1)
                Case 131
                    Call ModifyPKMNDouble(AtkPos, 2, 4, 1)
                'Dragon Dance (Atk +1, Spd +1)
                Case 133
                    Call ModifyPKMNDouble(AtkPos, 1, 7, 1)
                'Fake Tears (SDef -2)
                Case 138
                    Call ModifyPKMN(DefPos, 4, -2)
                'Flatter (SAtk +2, Confuse)
                Case 139
                    Call ModifyPKMN(DefPos, 3, 2)
                    If BCondition(DefPos).Substitute = 0 And BCondition(DefPos).Confuse = False Then
                        If Current(DefPos).Attribute = nbOwnTempo Then
                            Call AddMessage(145, nbActive, DefPos, nbTrait, nbOwnTempo)
                        ElseIf TeamBCondition(DefTeam).SafeguardCount <> 0 Then
                            Call AddMessage(79, nbActive, DefPos, nbMove, 173)
                        Else
                            BCondition(DefPos).Confuse = True
                            BCondition(DefPos).ConfuseCounter = EffectTook
                            Call AddMessage(193, nbActive, DefPos)
                        End If
                    End If
                'Follow Me
                Case 141
                    BCondition(AllyNum(AtkPos)).FollowMe = False
                    BCondition(AtkPos).FollowMe = True
                    Call AddMessage(258, nbActive, AtkPos)
                'Grudge
                Case 142
                    Call AddMessage(274, nbActive, AtkPos)
                    BCondition(AtkPos).Grudge = True
                'Hail
                Case 143
                    RealWeather = nbHailstorm
                    WeatherCount = 5
                    Call AddMessage(115)
                'Helping Hand
                Case 144
                    Call AddMessage(265, nbActive, AtkPos, nbActive, AtkAlly)
                    BCondition(AtkAlly).HelpingHand = True
                'Imprison
                Case 145
                    If BCondition(AtkPos).Imprison Then
                        Call AddMessage(38)
                    Else
                        Call AddMessage(273, nbActive, AtkPos)
                        BCondition(AtkPos).Imprison = True
                        For Y = 1 To 4
                            BCondition(AtkPos).SealedMove(Y) = Current(AtkPos).Move(Y)
                        Next Y
                    End If
                'Ingrain
                Case 146
                    If BCondition(AtkPos).Ingrain Then
                        Call AddMessage(38)
                    Else
                        Call AddMessage(88, nbActive, AtkPos)
                        BCondition(AtkPos).Ingrain = True
                    End If
                'Knock Off
                Case 147
                    If Current(DefPos).Item > nbNoItem Then
                        If Current(DefPos).Attribute = nbStickyHold Then
                            Call AddMessage(158, nbActive, DefPos, nbTrait, nbStickyHold, nbMove, M)
                        Else
                            Call AddMessage(271, nbActive, AtkPos, nbActive, DefPos, nbItem, Current(DefPos).Item)
                            Current(DefPos).Item = nbNoItem
                        End If
                    End If
                'Magic Coat
                Case 148
                    Call AddMessage(276, nbActive, AtkPos, nbMove, 302)
                    BCondition(AtkPos).MagicCoat = True
                'Memento (Atk -2, SAtk -2)
                Case 149
                    If BCondition(DefPos).Mist Or TeamBCondition(DefTeam).MistCount > 0 Then
                        Call AddMessage(96, nbActive, DefPos)
                    ElseIf Current(DefPos).Attribute = nbClearBody Or Current(DefPos).Attribute = nbWhiteSmoke Then
                        Call AddMessage(148, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
                    ElseIf Current(DefPos).Attribute = nbHyperCutter Then
                        Call AddMessage(149, nbActive, DefPos, nbTrait, Current(DefPos).Attribute, nbStat, 2)
                        Call ModifyPKMN(DefPos, 3, -2)
                    Else
                        Call ModifyPKMNDouble(DefPos, 1, 3, -2)
                    End If
                'Metal Sound (SDef -2)
                Case 150
                    Call ModifyPKMN(DefPos, 4, -2)
                'Mist Ball (SAtk -1)
                Case 151
                    Call ModifyPKMN(DefPos, 3, -1)
                'Mud Sport
                Case 152
                    If BCondition(AtkPos).MudSport Then
                        Call AddMessage(38)
                    Else
                        BCondition(AtkPos).MudSport = True
                        Call AddMessage(280)
                    End If
                'Overheat and Psycho Boost (User's SAtk -2)
                Case 155
                    Call ModifyPKMN(AtkPos, 3, -2)
                'Recycle
                Case 157
                    If Current(AtkPos).Item > nbNoItem Or Current(AtkPos).RecycleItem = nbNoItem Then
                        Call AddMessage(38)
                    Else
                        Call AddMessage(282, nbActive, AtkPos, nbItem, Current(AtkPos).RecycleItem)
                        Current(AtkPos).Item = Current(AtkPos).RecycleItem
                        Current(AtkPos).ItemEffect = EffectParam
                    End If
                'Refresh
                Case 158
                    Select Case Current(AtkPos).Condition
                    Case 2, 3, 5, 6
                        Call AddMessage(261, nbActive, AtkPos)
                        Current(AtkPos).Condition = 1
                        BCondition(AtkPos).ParalyzePenalty = False
                        BCondition(AtkPos).BurnPenalty = False
                    Case Else
                        Call AddMessage(38)
                    End Select
                'Roll Play
                Case 160
                    Call AddMessage(268, nbActive, AtkPos, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
                    Current(AtkPos).Attribute = Current(DefPos).Attribute
                'Skill Swap
                Case 162
                    Call AddMessage(272, nbActive, AtkPos, nbActive, DefPos)
                    Y = Current(AtkPos).Attribute
                    Current(AtkPos).Attribute = Current(DefPos).Attribute
                    Current(DefPos).Attribute = Y
                    BCondition(AtkPos).FlashFire = False
                    BCondition(DefPos).FlashFire = False
                'Smellingsalt
                Case 164
                    If Current(DefPos).Condition = 6 Then
                        Call AddMessage(190, nbActive, DefPos)
                        Current(DefPos).Condition = 1
                    End If
                'Snatch
                Case 165
                    Call AddMessage(278, nbActive, AtkPos)
                    For X = 1 To ActNum
                        For Y = 1 To ActNum
                            If BCondition(Y).Snatch = X Then Exit For
                        Next Y
                        If Y = ActNum + 1 Then
                            BCondition(AtkPos).Snatch = X
                            Exit For
                        End If
                    Next X
                'Stockpile
                Case 167
                    If BCondition(AtkPos).Stockpile = 3 Then
                        Call AddMessage(288, nbActive, AtkPos)
                    Else
                        BCondition(AtkPos).Stockpile = BCondition(AtkPos).Stockpile + 1
                        Call AddMessage(287, nbActive, AtkPos, nbNumber, BCondition(AtkPos).Stockpile)
                    End If
                'Superpower (User's Atk -1, User's Def -1)
                Case 168
                    Call ModifyPKMN(AtkPos, 1, -1)
                    Call ModifyPKMN(AtkPos, 2, -1)
                'Swallow
                Case 169
                    Y = BCondition(AtkPos).Stockpile
                    If Y = 0 Then
                        Call AddMessage(61)
                    ElseIf Current(AtkPos).HP = Current(AtkPos).MaxHP Then
                        Call AddMessage(81, nbActive, AtkPos)
                    Else
                        Call AddMessage(80, nbActive, AtkPos)
                        Select Case BCondition(AtkPos).Stockpile
                        Case 1: Call ChangeHP(AtkPos, 4)
                        Case 2: Call ChangeHP(AtkPos, 2)
                        Case 3: Call ChangeHP(AtkPos, 1)
                        End Select
                        BCondition(AtkPos).Stockpile = 0
                    End If
                'Taunt
                Case 170
                    If BCondition(DefPos).TauntCount > 0 Then
                        Call AddMessage(38)
                    Else
                        Call AddMessage(264, nbActive, DefPos)
                        BCondition(DefPos).TauntCount = 2
                    End If
                'Tickle (Atk -1, Def -1)
                Case 171
                    If BCondition(DefPos).Mist Or TeamBCondition(DefTeam).MistCount > 0 Then
                        Call AddMessage(96, nbActive, DefPos)
                    ElseIf Current(DefPos).Attribute = nbClearBody Or Current(DefPos).Attribute = nbWhiteSmoke Then
                        Call AddMessage(148, nbActive, DefPos, nbTrait, Current(DefPos).Attribute)
                    ElseIf Current(DefPos).Attribute = nbHyperCutter Then
                        Call AddMessage(148, nbActive, DefPos, nbTrait, Current(DefPos).Attribute, nbStat, 2)
                        Call ModifyPKMN(DefPos, 2, -1)
                    Else
                        Call ModifyPKMNDouble(DefPos, 1, 2, -1)
                    End If
                Case 172
                'Torment
                    If BCondition(DefPos).Torment Then
                        Call AddMessage(38)
                    Else
                        Call AddMessage(262, nbActive, DefPos)
                        BCondition(DefPos).Torment = True
                    End If
                'Trick
                Case 173
                    If Current(DefPos).Attribute = nbStickyHold Then
                        Call AddMessage(158, nbActive, DefPos, nbTrait, nbStickyHold, nbMove, M)
                    Else
                        X = Current(AtkPos).Item
                        Current(AtkPos).Item = Current(DefPos).Item
                        Current(DefPos).Item = X
                        X = Current(AtkPos).ItemEffect
                        Current(AtkPos).ItemEffect = Current(DefPos).ItemEffect
                        Current(DefPos).ItemEffect = X
                        BCondition(AtkPos).ChoiceBandSlot = 0
                        BCondition(DefPos).ChoiceBandSlot = 0
                        Call AddMessage(266, nbActive, AtkPos, nbActive, DefPos)
                        If Current(AtkPos).Item > nbNoItem Then Call AddMessage(267, nbActive, AtkPos, nbItem, Current(AtkPos).Item)
                        If Current(DefPos).Item > nbNoItem Then Call AddMessage(267, nbActive, DefPos, nbItem, Current(DefPos).Item)
                    End If
                'Uproar
                Case 174
                    'I've always wanted to wake someone up with a frying pan once...
                    If EffectTook <> 6 And BCondition(AtkPos).StuckMove = 0 Then
                        BCondition(AtkPos).StuckMove = M
                        BCondition(AtkPos).StuckCount = EffectTook
                        Call AddMessage(224, nbActive, AtkPos)
                    End If
                'Volt Tackle (1/3 recoil)
                Case 175
                    If Current(AtkPos).Attribute <> nbRockHead Then
                        X = Damage(1) \ 3
                        If X = 0 Or HitSubstitute Then X = 1
                        Current(AtkPos).HP = Current(AtkPos).HP - X
                        Call AddMessage(11, nbActive, AtkPos)
                    End If
                'Water Sport
                Case 176
                    Call AddMessage(281)
                    BCondition(AtkPos).WaterSport = True
                    If ActNum = 4 Then
                        If Current(AtkAlly).HP > 0 Then
                            BCondition(AtkAlly).WaterSport = True
                        End If
                    End If
                'Tail Glow (SAttack + 2)
                Case 177
                    Call ModifyPKMN(AtkPos, 3, 2)
                'Wish
                Case 179
                    Call AddMessage(86, nbActive, AtkPos)
                    BCondition(AtkPos).WishCount = 2
                    BCondition(AtkPos).WishPos = Current(AtkPos).TeamNumber
                'Yawn
                Case 180
                    Call AddMessage(270, nbActive, AtkPos, nbActive, DefPos)
                    BCondition(DefPos).YawnCount = 2
                    BCondition(DefPos).YawnSleepDuration = EffectTook
                End Select
            End If
            
            'King's Rock
            If KingsRock And Current(AtkPos).Item = nbKingsRock And BCondition(DefPos).Substitute <= 0 Then BCondition(DefPos).Flinching = True
            
            'FAINT CHECKS
            Call DoItems
            Call SyncPKMN(False, True, AtkPos)
            Call SyncPKMN(True, True, DefPos)
        End If
        
        'AFTER-DAMAGE TRAITS AND EFFECTS
        If Current(AtkPos).HP > 0 And Damage(Z) > 0 Then
            'Rage
            If BCondition(DefPos).RageCounter > 0 And Not HitSubstitute And TeamNum(DefPos) <> TeamNum(AtkPos) Then
                If BattleMode <> nbGSCBattle Then
                    If BCondition(AtkPos).AttackChange < 6 And AtkPos <> AllyNum(DefPos) Then
                        Call AddMessage(233, nbActive, DefPos)
                        Call ModifyPKMN(DefPos, 1, 1, True)
                    End If
                ElseIf BCondition(DefPos).RageCounter < 64 Then
                    Call AddMessage(233, nbActive, DefPos)
                    BCondition(DefPos).RageCounter = BCondition(DefPos).RageCounter + 1
                End If
            End If
            
            'Color Change
            If Current(DefPos).Attribute = nbColorChange And Current(DefPos).HP > 0 And Current(DefPos).Type1 <> TempMove.Type And Not HitSubstitute Then
                Call AddMessage(141, nbActive, DefPos, nbTrait, nbColorChange, nbType, TempMove.Type)
                Current(DefPos).Type1 = TempMove.Type
                Current(DefPos).Type2 = nbNoType
                TMATCH = TypeMatch(TempMove, DefPos)
            End If

            'Contact Traits
            If TraitEffect(Z) > 0 And TempMove.PhysMove And Not HitSubstitute Then
                Select Case Current(DefPos).Attribute
                Case nbStatic
                    If Current(AtkPos).Condition = 1 And Current(AtkPos).Attribute <> nbLimber And BCondition(AtkPos).Substitute = 0 Then
                        Call AddMessage(188, nbActive, DefPos, nbTrait, nbStatic, nbActive, AtkPos)
                        Current(AtkPos).Condition = 6
                        BCondition(AtkPos).ParalyzePenalty = True
                        Call DoSynchronize(DefPos, AtkPos)
                        Call DoItems
                    End If
                Case nbFlameBody
                    If Current(AtkPos).Condition = 1 And Current(AtkPos).Attribute <> nbWaterVeil And Not TypeCheck(Current(AtkPos), nbFire) And BCondition(AtkPos).Substitute = 0 Then
                        Call AddMessage(180, nbActive, DefPos, nbTrait, nbFlameBody, nbActive, AtkPos)
                        Current(AtkPos).Condition = 5
                        BCondition(AtkPos).BurnPenalty = True
                        Call DoSynchronize(DefPos, AtkPos)
                        Call DoItems
                    End If
                Case nbPoisonPoint
                    If Current(AtkPos).Condition = 1 And Current(AtkPos).Attribute <> nbImmunity And Not TypeCheck(Current(AtkPos), nbPoison, nbSteel) And BCondition(AtkPos).Substitute = 0 Then
                        Call AddMessage(175, nbActive, DefPos, nbTrait, nbPoisonPoint, nbActive, AtkPos)
                        Current(AtkPos).Condition = 2
                        Call DoSynchronize(DefPos, AtkPos)
                        Call DoItems
                    End If
                Case nbCuteCharm
                    If AttractOK(AtkPos, DefPos) And BCondition(AtkPos).Attract = 0 And Current(AtkPos).Attribute <> nbOblivious And Current(DefPos).HP > 0 Then
                        Call AddMessage(157, nbActive, DefPos, nbTrait, nbCuteCharm, nbActive, AtkPos)
                        BCondition(AtkPos).Attract = DefPos
                    End If
                Case nbEffectSpore
                    If Current(AtkPos).Condition = 1 And BCondition(AtkPos).Substitute = 0 Then
                        Select Case TraitEffect(Z)
                        Case 1
                            If Current(AtkPos).Attribute <> nbLimber Then
                                Call AddMessage(188, nbActive, DefPos, nbTrait, nbEffectSpore, nbActive, AtkPos)
                                Current(AtkPos).Condition = 6
                                BCondition(AtkPos).ParalyzePenalty = True
                                Call DoSynchronize(DefPos, AtkPos)
                            End If
                        Case 2
                            If Current(AtkPos).Attribute <> nbImmunity And Not TypeCheck(Current(AtkPos), nbPoison, nbSteel) Then
                                Call AddMessage(175, nbActive, DefPos, nbTrait, nbEffectSpore, nbActive, AtkPos)
                                Current(AtkPos).Condition = 2
                                Call DoSynchronize(DefPos, AtkPos)
                            End If
                        Case Else
                            If Current(AtkPos).Attribute <> nbInsomnia And Current(AtkPos).Attribute <> nbVitalSpirit Then
                                Call AddMessage(170, nbActive, DefPos, nbTrait, nbEffectSpore, nbActive, AtkPos)
                                Current(AtkPos).Condition = 4
                                Current(AtkPos).ConditionCount = TraitEffect(Z) - 1
                                Current(AtkPos).Resting = False
                                BCondition(AtkPos).Nightmare = False
                            End If
                        End Select
                        Call DoItems
                    End If
                Case nbRoughSkin
                    Call AddMessage(153, nbActive, DefPos, nbTrait, nbRoughSkin, nbActive, AtkPos)
                    Call ChangeHP(AtkPos, -16)
                End Select
            End If
                        
            Call SyncPKMN(False, True, AtkPos)
            If Multihitter And Current(AtkPos).HP <= 0 Then GoTo EndProcess
            
            'SHELL BELL
            If Current(AtkPos).Item = nbShellBell And Current(AtkPos).HP > 0 And Current(AtkPos).HP < Current(AtkPos).MaxHP Then
                X = Damage(Z) \ 8
                If X = 0 Then X = 1
                Call AddMessage(130, nbActive, AtkPos, nbItem, nbShellBell)
                Current(AtkPos).HP = Current(AtkPos).HP + X
            End If
        End If
        
        If MultihitExit Then
            HitCount = Z
            Exit For
        End If
        
    Next Z
       
    If Multihitter Then
        'Triple Kick is weird
        If TempMove.SpecialEffect = 118 And HitCount < 3 And BattleMode = nbAdvBattle Then Call AddMessage(37, nbActive, AtkPos)
        'Type Effectiveness message
        If Not SkipTypeMessage Then
            If TMATCH < 1 Then
                Call AddMessage(7)
            ElseIf TMATCH > 1 Then
                Call AddMessage(8)
            End If
        End If
        'HitCount message
        If TempMove.SpecialEffect <> 12 Then Call AddMessage(10, nbNumber, HitCount)
        'Damage message
        If BMessStyle = 1 And Not HitSubstitute Then Call DoDamageMessage(DefPos, MultihitDamage)
        
        'Jump to the effects block, then exit shortly after.
        If TempMove.SpecialEffect = 119 And EffectTook > 0 Then 'Twineedle's Poison
            Current(DefPos).Condition = 2
            Call AddMessage(174, nbActive, DefPos)
            Call DoSynchronize(AtkPos, DefPos)
            Call DoItems
        End If
        
        Call SyncPKMN(False, True, AtkPos)
        Call SyncPKMN(True, True, DefPos)
    End If
    
EndProcess:
    Call SyncPKMN(True)
    Exit Sub
DoTurnError:
    If InVBMode Then
        Stop
        Resume
    End If
End Sub
Private Function TypeMatch(TempMove As Move, ByVal DefPos As Byte, Optional ByVal Type1 As Elements = nbNoType, Optional ByVal Type2 As Elements = nbNoType) As Single
    'Figure out the Type matchup
    Dim TMATCH As Single
    If Type1 = nbNoType Then
        Type1 = Current(DefPos).Type1
        Type2 = Current(DefPos).Type2
    End If
    If Type2 > nbNoType And Not (BattleMode = nbRBYBattle And Type2 = 17) Then
        TMATCH = BattleMatrixEx(TempMove.Type, Type1, (BattleMode = nbRBYBattle)) * BattleMatrixEx(TempMove.Type, Type2, (BattleMode = nbRBYBattle))
    Else
        TMATCH = BattleMatrixEx(TempMove.Type, Type1, (BattleMode = nbRBYBattle))
    End If
    If BCondition(DefPos).Foresight And TMATCH = 0 Then
        If Type1 = 14 And Type2 = nbNoType Then TMATCH = 1
        If Type1 = 14 And Type2 > nbNoType Then TMATCH = BattleMatrix(TempMove.Type, Type2)
        If Type2 = 14 Then TMATCH = BattleMatrix(TempMove.Type, Type1)
    End If
    If TempMove.ID = 13 Then TMATCH = 1
    If TempMove.ID = 275 Then TMATCH = 1
    If TempMove.ID = 210 And BattleMode <> nbRBYBattle Then TMATCH = 1
    Select Case TempMove.SpecialEffect
    Case 76, 85, 103, 109
        If BattleMode = nbRBYBattle Or TMATCH > 0 Then TMATCH = 1
    End Select
    TypeMatch = TMATCH
End Function
Private Function UproarCheck(Pos) As Byte
    Dim X As Byte
    If Current(Pos).Attribute = nbCacophony Or Current(Pos).Attribute = nbSoundproof Then UproarCheck = 0: Exit Function
    For X = 1 To ActNum
        If BCondition(X).StuckMove = 346 Then UproarCheck = X: Exit Function
    Next X
    UproarCheck = 0
End Function
Private Function AirLocked() As Boolean
    Dim X As Byte
    For X = 1 To ActNum
        If (Current(X).Attribute = nbAirLock Or Current(X).Attribute = nbCloudNine) And Current(X).HP > 0 Then AirLocked = True: Exit Function
    Next X
    AirLocked = False
End Function
Public Sub SyncPKMN(Optional ByVal UseDB As Boolean = False, Optional SkipItems As Boolean = False, Optional ByVal FaintCheck As Byte = 0)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Byte
    Dim O() As Byte
    Dim EndTurnSync As Boolean
    EndTurnSync = Not (UseDB Or SkipItems)
    For Z = 1 To ActNum
        With BCondition(Z)
            Current(Z).HP = Int(Current(Z).HP)
            If Current(Z).HP > Current(Z).MaxHP Then Current(Z).HP = Current(Z).MaxHP
            .BideDamage = Cap(.BideDamage, 1023)
            If Current(Z).Condition <> 4 Then .Nightmare = False
            
            If Current(Z).HP <= 0 And Current(Z).Condition < 8 _
            And (Not BCondition(Z).FaintLast Or EndTurnSync) _
            And (FaintCheck = 0 Or FaintCheck = Z) Then
                If UseDB And .DBPos <> 0 And .DBPos <> AllyNum(Z) Then
                    If Current(.DBPos).HP > 0 Then
                        If .DestinyBond Then
                            Call AddMessage(241, nbActive, Z, nbActive, .DBPos)
                            Call SetFainted(.DBPos)
                        ElseIf .Grudge And BCondition(.DBPos).LastMoveSlot <> 0 Then
                            Call AddMessage(275, nbActive, .DBPos, nbMove, Current(.DBPos).Move(BCondition(.DBPos).LastMoveSlot))
                            Call DoPPChange(.DBPos, BCondition(.DBPos).LastMoveSlot, -Current(.DBPos).PP(BCondition(.DBPos).LastMoveSlot))
'                            Current(.DBPos).PP(BCondition(.DBPos).LastMoveSlot) = 0
                        End If
                    End If
                End If
                Call SetFainted(Z)
                
                If ActNum = 2 Then
                    SwitchTo(OtherTeam(Z)) = 0
                    Moved(OtherTeam(Z)) = False
                End If
            End If
            
            'RBY Hyper Beam no-recharge-on-faint
            If BattleMode = nbRBYBattle And (Current(1).HP = 0 Or Current(2).HP = 0) Then
                BCondition(1).Recharging = False
                BCondition(2).Recharging = False
            End If
        
            If Current(Z).No <> 0 Then
                X = TeamNum(Z)
                PKMN(X, Current(Z).TeamNumber).HP = Current(Z).HP
                PKMN(X, Current(Z).TeamNumber).Condition = Current(Z).Condition
                PKMN(X, Current(Z).TeamNumber).ConditionCount = Current(Z).ConditionCount
                PKMN(X, Current(Z).TeamNumber).Item = Current(Z).Item
                PKMN(X, Current(Z).TeamNumber).Resting = Current(Z).Resting
                PKMN(X, Current(Z).TeamNumber).RecycleItem = Current(Z).RecycleItem
            End If
            
            If EndTurnSync Then
                If .Soundproofed = 2 Then .LastMoveSlot = 0
                .Soundproofed = 0
            End If
        End With
    Next Z
    
    If BattleMode <> nbRBYBattle Then
        For X = 1 To ActNum
            If FaintCheck = X Or FaintCheck = 0 Then BCondition(X).DBPos = 0
            If BCondition(X).RepeatPos = 0 Then
                BCondition(X).RepeatMove = 0
                BCondition(X).RepeatCount = 0
            End If
        Next X
    End If
    
    For X = 1 To 6
        If PKMN(1, X).Condition <> 4 Then PKMN(1, X).Resting = False
        If PKMN(2, X).Condition <> 4 Then PKMN(2, X).Resting = False
    Next X
    
    Call DoForecast
    
    O = GetOrderArray
    'Switch Triggered Abilites
    
    'Weather Starters
    For Y = 1 To ActNum
        X = O(Y)
        If BCondition(X).SwitchedThisTurn And Current(X).HP > 0 Then
            If Current(X).Attribute = nbDrizzle And (RealWeather <> nbRaining Or WeatherCount <> 7) Then
                Call AddMessage(137, nbActive, X, nbTrait, nbDrizzle)
                RealWeather = nbRaining
                WeatherCount = 7
                Call DoForecast
            ElseIf Current(X).Attribute = nbDrought And (RealWeather <> nbSunny Or WeatherCount <> 7) Then
                Call AddMessage(138, nbActive, X, nbTrait, nbDrought)
                RealWeather = nbSunny
                WeatherCount = 7
                Call DoForecast
            ElseIf Current(X).Attribute = nbSandStream And (RealWeather <> nbSandstorm Or WeatherCount <> 7) Then
                Call AddMessage(139, nbActive, X, nbTrait, nbSandStream)
                RealWeather = nbSandstorm
                WeatherCount = 7
                Call DoForecast
            End If
        End If
    Next Y
    
    'Intimidate
    For Z = 1 To ActNum
        If BCondition(Z).SwitchedThisTurn And Current(Z).HP > 0 And Current(Z).Attribute = nbIntimidate Then
            For X = OtherTeam(TeamNum(Z)) To ActNum Step 2
                If Current(X).HP > 0 Then
                    If TeamBCondition(TeamNum(X)).MistCount > 0 Then
                        Call AddMessage(96, nbActive, X)
                    ElseIf Current(X).Attribute = nbClearBody Or Current(X).Attribute = nbWhiteSmoke Or Current(X).Attribute = nbHyperCutter Then
                        Call AddMessage(164, nbActive, X, nbTrait, Current(X).Attribute, nbActive, Z, nbTrait, nbIntimidate)
                    ElseIf BCondition(X).Substitute = 0 Then
                        Call AddMessage(152, nbActive, Z, nbTrait, nbIntimidate, nbActive, X)
                        Call ModifyPKMN(X, 1, -1, True)
                    End If
                End If
            Next X
        End If
    Next Z
           
    'Trace
    If Not NeedSwitch Then
        For Y = 1 To ActNum
            Z = O(Y)
            If Current(Z).HP > 0 And Current(Z).Attribute = nbTrace And BCondition(Z).TraceOK Then
                X = OtherTeam(TeamNum(Z))
                If ActNum = 4 Then
                    If (BCondition(Z).TraceNumber = 1 And Current(X + 2).HP > 0) Or Current(X).HP <= 0 Then X = X + 2
                End If
                Call AddMessage(154, nbActive, Z, nbActive, X, nbTrait, Current(X).Attribute)
                Current(Z).Attribute = Current(X).Attribute
            End If
            BCondition(Z).TraceOK = False
        Next Y
    End If
    
    For Y = 1 To ActNum
        If (Current(Y).Condition = 2 Or Current(Y).Condition = 3) And Current(Y).Attribute = nbImmunity Then
            Call AddMessage(159, nbActive, Y, nbTrait, nbImmunity)
            Current(Y).Condition = 1
        ElseIf Current(Y).Condition = 4 And (Current(Y).Attribute = nbInsomnia Or Current(Y).Attribute = nbVitalSpirit) Then
            Call AddMessage(299, nbActive, Y, nbTrait, Current(Y).Attribute)
            Current(Y).Condition = 1
        ElseIf Current(Y).Condition = 5 And Current(Y).Attribute = nbWaterVeil Then
            Call AddMessage(167, nbActive, Y, nbTrait, nbWaterVeil)
            Current(Y).Condition = 1
        ElseIf Current(Y).Condition = 6 And Current(Y).Attribute = nbLimber Then
            Call AddMessage(300, nbActive, Y, nbTrait, nbLimber)
            Current(Y).Condition = 1
        ElseIf Current(Y).Condition = 7 And Current(Y).Attribute = nbMagmaArmor Then
            Call AddMessage(301, nbActive, Y, nbTrait, nbMagmaArmor)
            Current(Y).Condition = 1
        ElseIf BCondition(Y).Confuse And Current(Y).Attribute = nbOwnTempo Then
            Call AddMessage(302, nbActive, Y, nbTrait, nbOwnTempo)
            BCondition(Y).Confuse = False
            BCondition(Y).ConfuseCounter = 0
        ElseIf BCondition(Y).Attract > 0 And Current(Y).Attribute = nbOblivious Then
            Call AddMessage(303, nbActive, Y, nbTrait, nbOblivious)
            BCondition(Y).Attract = 0
        End If
    Next Y
    
    For X = 1 To ActNum
        BCondition(X).SwitchedThisTurn = False
    Next X
    If Not SkipItems Then
        Call DoItems
        Call SyncPKMN(False, True, 5)
    End If
End Sub
Sub SetFainted(Pos As Byte)
    'Sets a Pokémon as Fainted, displays the "%1 fainted!" message,
    'and makes all other nessisary changes.  The PKMN variables are
    'synced early because that's the only way FaintedPokes will count
    'them.
    'With Doubles, this becomes a little more complex due to the fact
    'that, when only 1 poke remains, there will be a blank position
    'that NB is going to want to fill in.  That's what the BlankPos()
    'variable is for, and it's set here.  If you're wondering why
    'SwitchTo is checked here, it's because of a complicated situation
    'involving Spikes.
    Dim T As Byte
    Dim Y As Byte
    Dim BlankCondition As BattleStuff
    Call AddMessage(3, nbActive, Pos)
    T = TeamNum(Pos)
    Current(Pos).HP = 0
    Current(Pos).Condition = 8
    'Current(Pos).Attribute = nbNoTrait
    PKMN(T, Current(Pos).TeamNumber).HP = 0
    PKMN(T, Current(Pos).TeamNumber).Condition = 8
    Moved(Pos) = False
    For Y = 1 To ActNum
        If BCondition(Y).Attract = Pos Then BCondition(Y).Attract = 0
        If BCondition(Y).LockOn = Pos Then BCondition(Y).LockOn = 0
        If BCondition(Y).TrapNum = Pos Then BCondition(Y).TrapNum = 0
        If BCondition(Y).RepeatPos = Pos Then BCondition(Y).RepeatPos = 0
    Next Y

    With BCondition(Pos)
        'A few things carry over...
        BlankCondition.FutureSight = .FutureSight
        BlankCondition.TraceNumber = .TraceNumber
        BlankCondition.DoneCheck = .DoneCheck
        BlankCondition.WishCount = .WishCount
        BlankCondition.WishPos = .WishPos
    End With
    BCondition(Pos) = BlankCondition
    
    If ActNum = 4 Then
        If FaintedPokes(T) - (SwitchTo(AllyNum(Pos)) > 0) = NumPoke - 1 And BlankPos(T) = 0 Then
            BlankPos(T) = Pos
        End If
    End If
End Sub
Private Function TypeCheck(iPoke As Pokemon, ByVal CType1 As Elements, Optional ByVal CType2 As Elements = nbNoType, Optional ByVal CType3 As Elements = nbNoType) As Boolean
    TypeCheck = False
    If iPoke.Type1 = CType1 Or iPoke.Type2 = CType1 Then TypeCheck = True
    If CType2 > nbNoType And (iPoke.Type1 = CType2 Or iPoke.Type2 = CType2) Then TypeCheck = True
    If CType3 > nbNoType And (iPoke.Type1 = CType3 Or iPoke.Type2 = CType3) Then TypeCheck = True
End Function

Function AttractOK(AtkPos As Byte, DefPos As Byte) As Boolean
    With Me
        AttractOK = True
        If Current(AtkPos).Gender = 0 Then AttractOK = False
        If Current(DefPos).Gender = 0 Then AttractOK = False
        If Current(AtkPos).Gender = Current(DefPos).Gender Then AttractOK = False
    End With
End Function
Sub DoSynchronize(AtkPos As Byte, DefPos As Byte)
    If BCondition(AtkPos).Substitute > 0 Then Exit Sub
    Dim X As Integer
    If Current(DefPos).Attribute = nbSynchronize And Current(AtkPos).Condition = 1 Then
        Select Case Current(DefPos).Condition
        Case 2, 3
            If Current(AtkPos).Attribute = nbImmunity Then
                Call AddMessage(164, nbActive, AtkPos, nbTrait, nbImmunity, nbActive, DefPos, nbTrait, nbSynchronize)
            ElseIf TypeCheck(Current(AtkPos), nbPoison, nbSteel) Then
                Call AddMessage(168, nbActive, DefPos, nbTrait, nbSynchronize, nbActive, AtkPos)
            Else
                Current(AtkPos).Condition = 2
                Current(AtkPos).ConditionCount = 0
                Call AddMessage(175, nbActive, DefPos, nbTrait, nbSynchronize, nbActive, AtkPos)
            End If
        Case 5
            If Current(AtkPos).Attribute = nbWaterVeil Then
                Call AddMessage(164, nbActive, AtkPos, nbTrait, nbWaterVeil, nbActive, DefPos, nbTrait, nbSynchronize)
            ElseIf TypeCheck(Current(AtkPos), nbFire) Then
                Call AddMessage(168, nbActive, DefPos, nbTrait, nbSynchronize, nbActive, AtkPos)
            Else
                Current(AtkPos).Condition = 5
                Current(AtkPos).ConditionCount = 0
                BCondition(AtkPos).BurnPenalty = True
                Call AddMessage(180, nbActive, DefPos, nbTrait, nbSynchronize, nbActive, AtkPos)
            End If
        Case 6
            If Current(AtkPos).Attribute = nbLimber Then
                Call AddMessage(164, nbActive, AtkPos, nbTrait, nbLimber, nbActive, DefPos, nbTrait, nbSynchronize)
            Else
                Current(AtkPos).Condition = 6
                Current(AtkPos).ConditionCount = 0
                BCondition(AtkPos).ParalyzePenalty = True
                Call AddMessage(188, nbActive, DefPos, nbTrait, nbSynchronize, nbActive, AtkPos)
            End If
        End Select
    End If
End Sub
Function Odds(ByVal Percent As Single, Optional RandomNumber As Integer = -1) As Boolean
    'Figures odds using the Pokemon system
    'RBY maximum odds are 99.6%)
    Dim X As Integer
    Dim Y As Integer
    X = Round(Percent * 256 / 100, 0) - 1
    If RandomNumber = -1 Then Y = GetNextRnd Else Y = RandomNumber
    If BattleMode = nbRBYBattle Then Y = Y + 1
    If Y <= X Then Odds = True Else Odds = False
    'Odds = False
    'Odds = True
End Function
Private Function TestOdds()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    For X = 1 To 8000
        If Odds(50) Then Y = Y + 1 Else Z = Z + 1
    Next X
    'Debug.Print Y & " True, " & Z & " Fasle."
End Function
Function GetPresentValue(ByVal PKMNType As Byte) As Byte
    Select Case PKMNType
        Case 1
            GetPresentValue = 0
        Case 2
            GetPresentValue = 20
        Case 3
            GetPresentValue = 21
        Case 4
            GetPresentValue = 23
        Case 5
            GetPresentValue = 22
        Case 6
            GetPresentValue = 25
        Case 7
            GetPresentValue = 1
        Case 8
            GetPresentValue = 3
        Case 9
            GetPresentValue = 4
        Case 10
            GetPresentValue = 2
        Case 11
            GetPresentValue = 24
        Case 12
            GetPresentValue = 7
        Case 13
            GetPresentValue = 5
        Case 14
            GetPresentValue = 8
        Case 15
            GetPresentValue = 26
        Case 16
            GetPresentValue = 27
        Case 17
            GetPresentValue = 9
    End Select
End Function
Public Function GetSPEffect()
    Select Case Terrain
    Case nbShortGrass, nbTallGrass
        GetSPEffect = 81 'Poison
    Case nbVeryTallGrass
        GetSPEffect = 64 'Sleep
    Case nbOcean
        GetSPEffect = 9 'Attack - 1
    Case nbUnderwater
        GetSPEffect = 2 'Defense - 1
    Case nbPond
        GetSPEffect = 20 'Speed - 1
    Case nbsAnd
        GetSPEffect = 50 'Accuracy - 1
    Case nbCave
        GetSPEffect = 16 'Flinch
    Case nbMountain
        GetSPEffect = 22 'Confusion
    Case Else
        GetSPEffect = 18 'Paralyze
    End Select
End Function
Public Function AccuracyMod(ByVal Ratio As Single) As Single
    AccuracyMod = ((Abs(Ratio) + 3) / 3) ^ Sgn(Ratio)
End Function
Public Function Rollover(ByVal Value As Integer) As Integer
    If StadiumMode Or BattleMode = nbAdvBattle Then
        Rollover = Value
        Exit Function
    End If
    If Value <= 1024 Then
        Rollover = Value
    ElseIf Value > 1024 And Value <= 2048 Then
        Rollover = Value - 1024
    Else
        Rollover = Value - 2048
    End If
End Function
Private Sub DoTransform(ByVal AtkPos As Byte, ByVal Target As Byte)

End Sub


Sub ModifyPKMN(ByVal Pos As Byte, ByVal Stat As Long, ByVal Change As Long, Optional SkipMessage As Boolean = False)
    Dim TCWak As Boolean
    Dim LBPika As Boolean
    Dim MPDitto As Boolean
    Dim X As Long
    Dim Z As Long
    Dim Temp As Long
    Dim StartValue As Long
    Dim UseText As Long
    Dim MaxChange As Long
    Dim SC As Long
        
    'Exit if opponent is KOed.
    If Current(Pos).HP <= 0 And Current(Pos).Item <> nbBerserkGene Then Exit Sub
    If BattleMode = nbRBYBattle And (Stat = 3 Or Stat = 4) Then Stat = 8
    
    'Set the initial values.
    TCWak = ((Current(Pos).No = 104 Or Current(Pos).No = 105) And Current(Pos).Item = nbThickClub)
    LBPika = (Current(Pos).No = 25 And Current(Pos).Item = nbLightBall)
    MPDitto = (PKMN(TeamNum(Pos), Current(Pos).TeamNumber).No = 132 And Current(Pos).Item = nbMetalPowder)
    Select Case Stat
    Case 1
        SC = BCondition(Pos).AttackChange
        Z = Current(Pos).Attack
        If TCWak Then Z = Z * 2
        UseText = 2
    Case 2
        SC = BCondition(Pos).DefenseChange
        Z = Current(Pos).Defense
        If MPDitto Then Z = Z * 1.5
        UseText = 3
    Case 3
        SC = BCondition(Pos).SAttackChange
        Z = Current(Pos).SpecialAttack
        If LBPika Then Z = Z * 2
        UseText = 6
    Case 4
        SC = BCondition(Pos).SDefenseChange
        Z = Current(Pos).SpecialDefense
        If MPDitto Then Z = Z * 1.5
        UseText = 7
    Case 5
        SC = BCondition(Pos).AccuracyChange
        Z = 0
        UseText = 8
    Case 6
        SC = BCondition(Pos).EvadeChange
        Z = 0
        UseText = 9
        'If BCondition(pos).Foresight And Change > 0 Then Change = 0
    Case 7
        SC = BCondition(Pos).SpeedChange
        Z = Current(Pos).Speed
        UseText = 4
    Case 8
        SC = BCondition(Pos).SAttackChange
        Z = Current(Pos).SpecialAttack
        UseText = 5
    End Select
    
    'Figure out the max
    For X = 1 To 6
        If Z * StatChange(X) >= 999 Then Exit For
    Next X
    If X = 7 Or BattleMode = nbAdvBattle Then X = 6
    MaxChange = X

    'Fail Messages
    If Not SkipMessage Then
        If SC = MaxChange And Sgn(Change) = 1 Then
            Call AddMessage(63, nbActive, Pos, nbStat, UseText): Exit Sub
        ElseIf SC = -6 And Sgn(Change) = -1 Then
            Call AddMessage(64, nbActive, Pos, nbStat, UseText): Exit Sub
        End If
    End If
    
    'Do the change
    SC = SC + Change
    If SC < -6 Then SC = -6
    If SC > MaxChange Then SC = MaxChange
    
    'Update the actual variables
    With BCondition(Pos)
        Select Case Stat
        Case 1: .AttackChange = SC
        Case 2: .DefenseChange = SC
        Case 3: .SAttackChange = SC
        Case 4: .SDefenseChange = SC
        Case 5: .AccuracyChange = SC
        Case 6: .EvadeChange = SC
        Case 7: .SpeedChange = SC
        Case 8
            .SAttackChange = SC
            .SDefenseChange = SC
        End Select
    End With
    
    'RBY Mode Burn/Par Glitch
    If BattleMode = nbRBYBattle Then
        Select Case Stat
            Case 1
                'For RBY Burns:
                If Change <> 0 And Current(Pos).Condition = 5 Then
                    BCondition(Pos).BurnPenalty = (Sgn(Change) <> 1)
                End If
            Case 7
                'For RBY Paralysis:
                If Change <> 0 And Current(Pos).Condition = 6 Then
                    BCondition(Pos).ParalyzePenalty = (Sgn(Change) <> 1)
                End If
        End Select
    End If
    
    'Do the message
    If Not SkipMessage Then
        'Do the animation
        X = 1 * 2
        If Change < 0 Then X = X + 1
        'Call DoAnim(BattleWindow, nbaStatus, Pos, X)
    
        Select Case Change
            Case -2
                Call AddMessage(16, nbActive, Pos, nbStat, UseText)
            Case -1
                Call AddMessage(14, nbActive, Pos, nbStat, UseText)
            Case 1
                Call AddMessage(13, nbActive, Pos, nbStat, UseText)
            Case 2
                Call AddMessage(15, nbActive, Pos, nbStat, UseText)
        End Select
    End If
End Sub
Sub ModifyPKMNDouble(Poke As Byte, Stat1 As Integer, Stat2 As Integer, Change As Integer, Optional SkipFailMessage As Boolean = False)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Integer
    Dim B As Integer
    For X = 1 To 2
        If X = 1 Then Y = Stat1 Else Y = Stat2
        With BCondition(Poke)
            Select Case Y
            Case 1: B = .AttackChange
            Case 2: B = .DefenseChange
            Case 3: B = .SAttackChange
            Case 4: B = .SDefenseChange
            Case 5: B = .AccuracyChange
            Case 6: B = .EvadeChange
            Case 7: B = .SpeedChange
            End Select
        End With
        If X = 1 Then A = B
    Next X
    X = 0
    If A <> 6 * Sgn(Change) Then
        X = 1
        Call ModifyPKMN(Poke, Stat1, Change)
    End If
    If B <> 6 * Sgn(Change) Then
        X = 1
        Call ModifyPKMN(Poke, Stat2, Change)
    End If
    If X = 0 And Not SkipFailMessage Then Call AddMessage(IIf(Sgn(Change) = 1, 65, 66), nbActive, Poke)
End Sub

Sub DoSwitch(ByVal Team As Byte, ByVal Number As Integer, Optional ByVal BatonPass As Boolean = False, Optional SkipSync As Boolean = False, Optional SkipMessage As Boolean = False)
    Dim BlankCondition As BattleStuff
    Dim BlankTC As TeamCond
    Dim X As Integer
    Dim Y As Byte
    Dim B As Boolean
    X = Current(Team).TeamNumber
    B = NeedSwitch(Team)
    If Not (BatonPass Or SkipMessage Or BCondition(Team).BeingPursued) Then
        If Current(Team).HP > 0 And Current(Team).TeamNumber <> Number Then
            Call AddMessage(297, nbPlayer, TeamNum(Team), nbActive, Team)
        End If
    End If
    
    'If Not BatonPass Then Call DoAnim(BattleWindow, nbaSwitch, Team, Int(Rnd * 12))
    
    If Current(Team).Attribute = nbNaturalCure And Current(Team).Condition < 8 Then PKMN(TeamNum(Team), Current(Team).TeamNumber).Condition = 1
    For Y = 1 To ActNum
        If BCondition(Y).LockOn = Team And Not BatonPass Then BCondition(Y).LockOn = 0
        If BCondition(Y).Attract = Team Then BCondition(Y).Attract = 0
        If BCondition(Y).RepeatPos = Team Then BCondition(Y).RepeatPos = 0
        If BCondition(Y).TrapNum = Team And Not BatonPass Then BCondition(Y).TrapNum = 0
    Next Y
    If BattleMode = nbRBYBattle Then
        With BCondition(OtherTeam(Team))
            If .RepeatCount > 0 Then
                .RepeatCount = 0
                .RepeatMove = 0
            End If
        End With
    End If
    If BatonPass Then
        BlankCondition = BCondition(Team)
        With BlankCondition
            .Attract = 0
            .DestinyBond = False
            .FuryCutter = 0
            .MimicedMove = 0
            .ProtectPercent = 0
            .RageCounter = 0
            .Rollout = 0
            .TransformedTo = 0
            .LastMove = 0
            .LastMoveSlot = 0
            If BattleMode = nbAdvBattle Then
                .YawnCount = 0
                .YawnSleepDuration = 0
                .ChoiceBandSlot = 0
                .Imprison = False
                .Torment = False
                .TauntCount = 0
                .DefenseCurl = False
                .Minimize = False
                .Foresight = False
                .Nightmare = False
                .RepeatCount = 0
                .RepeatMove = 0
                .RepeatPos = 0
            Else
                If PKMN(Team, Number).Condition = 2 And Current(Team).Condition = 3 Then
                    PKMN(Team, Number).Condition = 3
                    .ToxicCount = 0
                End If
                If PKMN(Team, Number).Condition <> 4 Then .Nightmare = False
            End If
        End With
    End If
    
    'A few things carry over...
    BlankCondition.FutureSight = BCondition(Team).FutureSight
    BlankCondition.TraceNumber = BCondition(Team).TraceNumber
    BlankCondition.DoneCheck = BCondition(Team).DoneCheck
    BlankCondition.WishCount = BCondition(Team).WishCount
    BlankCondition.WishPos = BCondition(Team).WishPos
    
    'RBY Reflect and Light Screen fade
    If BattleMode = nbRBYBattle Then
        TeamBCondition(Team).ReflectCount = 0
        TeamBCondition(Team).LightScreenCount = 0
    End If
    
    If BattleMode <> nbAdvBattle Then   'Toxic becomes regular poison
        If PKMN(TeamNum(Team), Current(Team).TeamNumber).Condition = 3 Then PKMN(TeamNum(Team), Current(Team).TeamNumber).Condition = 2
    End If
    
    BCondition(Team) = BlankCondition
    If BattleMode = nbRBYBattle Then TeamBCondition(Team) = BlankTC
    Current(Team) = PKMN(TeamNum(Team), Number)
    Current(Team).TeamNumber = Number
    BCondition(Team).BurnPenalty = (Current(Team).Condition = 5)
    BCondition(Team).ParalyzePenalty = (Current(Team).Condition = 6)
    If BatonPass And Not StadiumMode And BattleMode <> nbAdvBattle Then
        If PKMN(TeamNum(Team), X).Condition = 5 And Current(Team).Condition = 5 Then
            BCondition(Team).BurnPenalty = False
        End If
        If PKMN(TeamNum(Team), X).Condition = 6 And Current(Team).Condition = 6 Then
            BCondition(Team).ParalyzePenalty = False
        End If
    End If
    BCondition(Team).SwitchedThisTurn = True
    BCondition(Team).TraceOK = True
    BCondition(Team).FakeOut = 2
    If Current(Team).Attribute = nbTruant Then BCondition(Team).Truant = True
    
    'If Not BatonPass Then Call DoAnim(BattleWindow, nbaSwitch, Team, Int(Rnd * 12))
    
    If Not SkipMessage Then
        Call AddMessage(298, nbPlayer, TeamNum(Team), nbActive, Team, nbNumber, Current(Team).Level, nbPoke, Current(Team).No)
    End If
    
    'Call DoAnim(BattleWindow, nbaSwitch + 1, Team, Int(Rnd * 12), ActNum)
    
    If TeamBCondition(TeamNum(Team)).Spikes > 0 Then
        If Not TypeCheck(Current(Team), nbFlying) And Current(Team).Attribute <> nbLevitate Then
            Call AddMessage(246, nbActive, Team)
            X = Current(Team).MaxHP * (TeamBCondition(TeamNum(Team)).Spikes + 1) / 16
            If X = 0 Then X = 1
            Current(Team).HP = Current(Team).HP - X
        End If
    End If
    If BatonPass Then BCondition(Team).BatonPassing = False
    Call SyncPKMN
    Call DoBerserkGene(Team)
    If BattleMode <> nbAdvBattle Then
        If BCondition(OtherTeam(Team)).InPursuit Or Not B Then Call DoConditionsGSC(Team)
    End If
    SwitchedThisTurn(Team) = True
    SwitchTo(Team) = 0
    Moved(Team) = False
End Sub
Sub DoStartSwitch(Team As Byte)
    Dim X As Byte
    Dim E() As Byte
    Current(Team) = PKMN(Team, 1)
    Current(Team).TeamNumber = 1
    Call AddMessage(298, nbPlayer, Team, nbActive, Team, nbNumber, Current(Team).Level, nbPoke, Current(Team).No)
    If ActNum = 4 Then
        Current(Team + 2) = PKMN(Team, 2)
        Current(Team + 2).TeamNumber = 2
        Call AddMessage(298, nbPlayer, Team, nbActive, Team + 2, nbNumber, Current(Team + 2).Level, nbPoke, Current(Team + 2).No)
    End If
    Call DoBerserkGene(Team)
    For X = 1 To ActNum
        BCondition(X).FakeOut = 2
        BCondition(X).TraceOK = True
        BCondition(X).SwitchedThisTurn = True
    Next X
End Sub
Private Sub DoBerserkGene(ByVal Pos As Byte)
    Dim E() As Byte
    Dim X As Long
    If Current(Pos).Item = nbBerserkGene Then
        Call AddMessage(125, nbActive, Pos, nbItem, 12)
        Call ModifyPKMN(Pos, 1, 2)
        Current(Pos).Item = nbNoItem
        PKMN(Pos, Current(Pos).TeamNumber).Item = nbNoItem
        BCondition(Pos).Confuse = True
        If EnabledRule(nbStadiumMode) Then
            E = ItemEffect2Array(Current(Pos).ItemEffect)
            X = E(1)
            If X = 5 Then X = E(2)
            BCondition(Pos).ConfuseCounter = X
        Else
            BCondition(Pos).ConfuseCounter = 5
        End If
        Call AddMessage(193, nbActive, Pos)
    End If
End Sub
Sub DoForecast()
    Dim X As Long
    Dim Y As Elements
    For X = 1 To ActNum
        If Current(X).No = 351 And Current(X).Attribute = nbForecast And Current(X).HP > 0 Then
            Select Case CurrentWeather
            Case nbNoWeather, nbSandstorm
                Y = nbNormal
            Case nbRaining
                Y = nbWater
            Case nbSunny
                Y = nbFire
            Case nbHailstorm
                Y = nbIce
            End Select
            If Current(X).Type1 <> Y Then
                Call AddMessage(161, nbActive, X)
                Current(X).Type1 = Y
            End If
        End If
    Next X
End Sub
Function RuleCheck(ByVal Team As Integer, ByVal Clause As Integer, Optional ByVal Condition As Integer = 0) As Boolean
    Dim X As Byte
    Dim ClauseCount As Byte
        Select Case Clause
            Case 1
                If Not EnabledRule(nbSleep) Then
                    RuleCheck = False
                Else
                    For X = 1 To NumPoke
                        If PKMN(Team, X).Condition = 4 And Not PKMN(Team, X).Resting Then
                            ClauseCount = ClauseCount + 1
                        End If
                    Next
                    RuleCheck = (ClauseCount > 0)
                End If
            Case 2
                'RuleCheck = EnabledRule()
            Case 3
                If Not EnabledRule(nbFreeze) Then
                    RuleCheck = False
                Else
                    For X = 1 To NumPoke
                        If PKMN(Team, X).Condition = 7 Then ClauseCount = ClauseCount + 1
                    Next
                    RuleCheck = (ClauseCount > 0)
                End If
        End Select
End Function

Sub DoCountdown(ByVal Team As Byte)

    'Count down a condition
    'RBY is handled elsewhere
    If Not BattleMode = nbRBYBattle Then
        If Current(Team).ConditionCount > 0 Then Current(Team).ConditionCount = Current(Team).ConditionCount - 1
        If Current(Team).Condition = 4 And Current(Team).Attribute = nbEarlyBird And Current(Team).ConditionCount > 0 Then
            Current(Team).ConditionCount = Current(Team).ConditionCount - 1
        End If
        'Handle acondition wearing off
        If Current(Team).ConditionCount = 0 Then
            Select Case Current(Team).Condition
                Case 4
                    Call AddMessage(173, nbActive, Team)
                    Current(Team).Condition = 1
                    Current(Team).Resting = False
                    BCondition(Team).Nightmare = False
                    PKMN(TeamNum(Team), Current(Team).TeamNumber).Resting = False
            End Select
        End If
    End If
    
    If BattleMode <> nbAdvBattle Then
        'Confusion
        If BCondition(Team).ConfuseCounter = 0 And BCondition(Team).Confuse Then
            BCondition(Team).Confuse = False
            Call AddMessage(192, nbActive, Team)
        End If
        If BCondition(Team).ConfuseCounter > 0 And BCondition(Team).ConfuseCounter < 5 Then
            BCondition(Team).ConfuseCounter = BCondition(Team).ConfuseCounter - 1
        End If
    
        'Disable
        If BCondition(Team).DisabledMove > 0 Then
            BCondition(Team).DisableCount = BCondition(Team).DisableCount - 1
            If BCondition(Team).DisableCount = 0 Then
                Call AddMessage(235, nbActive, Team)
                BCondition(Team).DisabledMove = 0
            End If
        End If
    End If
    
    PKMN(TeamNum(Team), Current(Team).TeamNumber).ConditionCount = Current(Team).ConditionCount
    PKMN(TeamNum(Team), Current(Team).TeamNumber).Condition = Current(Team).Condition
End Sub
Function Shorten(ByVal Number As Long, ByVal Slot As Long) As Long
    If Number > Slot Then Shorten = Number - 1 Else Shorten = Number
End Function
Function Expand(ByVal Number As Long, ByVal Slot As Long) As Long
    If Number >= Slot Then Expand = Number + 1 Else Expand = Number
End Function
Function LastPKMN(ByVal PNum As Byte)
    Dim X As Byte
    Dim F As Byte
    
    For X = 1 To NumPoke
        If PKMN(PNum, X).Condition = 8 Then F = F + 1
    Next
    If F = NumPoke - 1 Then LastPKMN = True Else LastPKMN = False
End Function

Public Property Get VariableSync() As String
    Dim Build As String
    Dim Temp As String
    Dim X As Integer
    Dim Y As Integer
    Dim d As Long
    Dim PNum As Byte
    On Error GoTo ETrap
    
    SetDestStringASM VarPtr(Build), d
    StreamInASM BattleMode, 2
    StreamInASM Abs(ActNum = 4), 1
    For X = UBound(RealRule) To 1 Step -1
        StreamInASM Abs(RealRule(X)), 1
    Next X
    StreamInASM Abs(WatchIgnore1), 1
    StreamInASM Abs(WatchIgnore2), 1
    StreamInASM NumPoke, 3

    'First the Pokes
    For PNum = 1 To 2
        StreamInASM GameVersion(PNum), 3
        StreamInASM PKMN(PNum, 1).GameVersion, 3
        For Y = 1 To NumPoke
            With PKMN(PNum, Y)
                X = .Gender - 1
                If X = -1 Then X = 0
                StreamInASM .No, 9
                StreamInASM .Level, 7
                StreamInASM .Item, 7
                StreamInASM .NatureNum, 5
                StreamInASM .AttNum, 1
                StreamInASM X, 1
                StreamInASM Abs(.Shiny), 1
                StreamInASM .UnownLetter, 5
                For X = 1 To 4
                    StreamInASM .Move(X), 9
                    StreamInASM .PP(X), 7
                Next X
                StreamInASM .DV_HP, 5
                StreamInASM .DV_Atk, 5
                StreamInASM .DV_Def, 5
                StreamInASM .DV_Spd, 5
                StreamInASM .DV_SAtk, 5
                StreamInASM .DV_SDef, 5
                StreamInASM .EV_HP, 8
                StreamInASM .EV_Atk, 8
                StreamInASM .EV_Def, 8
                StreamInASM .EV_Spd, 8
                StreamInASM .EV_SAtk, 8
                StreamInASM .EV_SDef, 8
                StreamInASM .Condition, 4
                StreamInASM .ConditionCount, 4
                StreamInASM IIf(.HP < 0, 0, .HP), BinDig(.MaxHP)
                StreamInASM .ItemEffect, 8
                StreamInASM .RecycleItem, 7
                StreamInASM Abs(.Resting), 1
                Temp = .Nickname
                If .Nickname = .Name Then Temp = ""
                StreamInASM Len(Temp), 4
                For X = 1 To Len(Temp)
                    StreamInASM Asc(Mid$(Temp, X, 1)), 8
                Next X
            End With
        Next Y
    Next PNum
    
    'Battle Variables
    StreamInASM RealWeather, 3
    StreamInASM WeatherCount, 3
    StreamInASM TurnNumber, 10
    StreamInASM Terrain, 4
    StreamInASM BattlePauseNum, 3
    For X = 1 To ActNum
        StreamInASM Order(X).MoveRandom, 3
        StreamInASM Order(X).EndTurnRandom, 3
        StreamInASM Order(X).PriorityTier + 5, 4
        StreamInASM Abs(Unfreezing(X)), 1
        StreamInASM Abs(SwitchedThisTurn(X)), 1
    Next X
    For X = 1 To 2
        StreamInASM BlankPos(X), 3
        StreamInASM Abs(SelfLoss(X)), 1
    Next X
    
    'TransformedTo need higher priority
    For X = 1 To ActNum
        StreamInASM Current(X).TeamNumber, 3
        StreamInASM BCondition(X).TransformedTo, 4
    Next X
    
    'and finally, the current conditions
    For X = 1 To ActNum
        With Current(X)
            StreamInASM .Type1, 5
            StreamInASM .Type2, 5
            StreamInASM .Attribute, 7
        End With
        With BCondition(X)
            StreamInASM .AttackChange + 6, 4
            StreamInASM .DefenseChange + 6, 4
            StreamInASM .SpeedChange + 6, 4
            StreamInASM .SAttackChange + 6, 4
            StreamInASM .SDefenseChange + 6, 4
            StreamInASM .AccuracyChange + 6, 4
            StreamInASM .EvadeChange + 6, 4
            StreamInASM .StuckMove, 9
            StreamInASM .StuckCount, 3
            StreamInASM .RepeatMove, 9
            StreamInASM .RepeatCount, 3
            StreamInASM .DisabledMove, 3
            StreamInASM .DisableCount, 4
            StreamInASM .EncoreMove, 3
            StreamInASM .EncoreDuration, 3
            StreamInASM .BideDamage, 10
            StreamInASM .BideCount, 2
            StreamInASM .LockOnCount, 2
            StreamInASM .ConfuseCounter + 1, 4
            StreamInASM Shorten(.Attract, X), 2
            StreamInASM Shorten(.LeechSeed, X), 2
            StreamInASM Shorten(.TrapNum, X), 2
            StreamInASM Shorten(.LockOn, X), 2
            StreamInASM Shorten(.RepeatPos, X), 2
            StreamInASM .ToxicCount, 4
            StreamInASM .PerishSong, 3
            StreamInASM .Rollout, 3
            StreamInASM .FuryCutter, 3
            StreamInASM .ProtectPercent, 7
            StreamInASM .RageCounter, 6
            StreamInASM .MimicedMove, 9
            StreamInASM .Substitute, 10
            StreamInASM .LastMoveSlot, 3
            StreamInASM .Stockpile, 2
            StreamInASM .YawnCount, 2
            StreamInASM .YawnSleepDuration, 4
            StreamInASM .ChargeCount, 2
            StreamInASM .WishCount, 2
            StreamInASM .WishPos, 3
            StreamInASM .SemiInvul, 3
            StreamInASM .Snatch, 3
            StreamInASM .TraceNumber, 1
            StreamInASM .OutrageConfuse, 4
            StreamInASM .TauntCount, 3
            StreamInASM .ChoiceBandSlot, 3
            StreamInASM .FakeOut, 2
            StreamInASM Abs(.Confuse), 1
            StreamInASM Abs(.Charging), 1
            StreamInASM Abs(.Recharging), 1
            StreamInASM Abs(.FocusEnergy), 1
            StreamInASM Abs(.Foresight), 1
            StreamInASM Abs(.DestinyBond), 1
            StreamInASM Abs(.Nightmare), 1
            StreamInASM Abs(.DefenseCurl), 1
            StreamInASM Abs(.Curse), 1
            StreamInASM Abs(.Minimize), 1
            StreamInASM Abs(.Mist), 1
            StreamInASM Abs(.Encore), 1
            StreamInASM Abs(.BurnPenalty), 1
            StreamInASM Abs(.ParalyzePenalty), 1
            StreamInASM Abs(.BatonPassing), 1
            StreamInASM Abs(.InPursuit), 1
            StreamInASM Abs(.BeingPursued), 1
            StreamInASM Abs(.Flinching), 1
            StreamInASM Abs(.Protected), 1
            StreamInASM Abs(.Enduring), 1
            StreamInASM Abs(.Damaged), 1
            StreamInASM Abs(.HelpingHand), 1
            StreamInASM Abs(.Grudge), 1
            StreamInASM Abs(.Ingrain), 1
            StreamInASM Abs(.Truant), 1
            StreamInASM Abs(.ShedSkin), 1
            StreamInASM Abs(.MagicCoat), 1
            StreamInASM Abs(.FlashFire), 1
            StreamInASM Abs(.DoneCheck), 1
            StreamInASM Abs(.SwitchedThisTurn), 1
            StreamInASM Abs(.LansatBerry), 1
            StreamInASM Abs(.WaterSport), 1
            StreamInASM Abs(.MudSport), 1
            StreamInASM Abs(.Torment), 1
            StreamInASM Abs(.FaintLast), 1
            StreamInASM Abs(.TraceOK), 1
            StreamInASM Abs(.FollowMe), 1
            StreamInASM Abs(.Imprison), 1
            If .Imprison Then
                For Y = 1 To 4
                    StreamInASM .SealedMove(Y), 9
                Next Y
            End If
            If .FutureSight.count = 0 Then
                StreamInASM 0, 1
            Else
                StreamInASM 1, 1
                StreamInASM .FutureSight.AccuracyMod + 6, 4
                StreamInASM .FutureSight.AttackPower, 13
                StreamInASM Abs(.FutureSight.CHit), 1
                StreamInASM .FutureSight.count, 3
                StreamInASM .FutureSight.FSMove, 1
                StreamInASM .FutureSight.Level, 7
                StreamInASM .FutureSight.RandomNumber, 8
            End If
        End With
    Next X
    For X = 1 To 2
        With TeamBCondition(X)
            StreamInASM .Spikes, 2
            StreamInASM .ReflectCount, 3
            StreamInASM .LightScreenCount, 3
            StreamInASM .SafeguardCount, 3
            StreamInASM .MistCount, 3
        End With
    Next X
    ''Debug.Print "Full Sync Binary Len: " & Len(Build)
    'Build = Bin2Chr(Build)
    ''Debug.Print "Full Sync Converted Len: " & Len(Build)
    VariableSync = Build
    Exit Property
ETrap:
    Stop
    Resume
End Property
Public Property Let VariableSync(SyncString As String)
    Dim X As Integer
    Dim Y As Integer
    Dim Nick As String
    Dim PNum As Byte
    Dim BlankPKMN As Pokemon
    Dim BuildPKMN As Pokemon
    Dim Build As String
    Dim GV As Long
    Build = Chr2Bin(SyncString)
    Call ParseString(0, Build)
    BattleMode = Bin2Dec(ParseString(2))
    ActNum = IIf(ParseString(1) = "0", 2, 4)
    Me.Rules = FixedHex(Bin2Dec(ParseString(UBound(RuleText))), 4)
    WatchIgnore1 = Bin2Bool(ParseString(1))
    WatchIgnore2 = Bin2Bool(ParseString(1))
    NumPoke = Bin2Dec(ParseString(3))
    If DoMessages Then
        If WatchIgnore1 Then
            Call LogOutput.AddMessage(Player(Player1).Name & " is ignoring spectator chat.", , 38400, True)
        End If
        If WatchIgnore2 Then
            Call LogOutput.AddMessage(Player(Player2).Name & " is ignoring spectator chat.", , 38400, True)
        End If
    End If
    For PNum = 1 To 2
        GameVersion(PNum) = Bin2Dec(ParseString(3))
        GV = Bin2Dec(ParseString(3))
        For X = 1 To NumPoke
            PKMN(PNum, X) = BasePKMN(Bin2Dec(ParseString(9)))
            With PKMN(PNum, X)
                .GameVersion = GV
                .Level = Bin2Dec(ParseString(7))
                .Item = Bin2Dec(ParseString(7))
                .NatureNum = Bin2Dec(ParseString(5))
                .AttNum = Val(ParseString(1))
                .Gender = Bin2Dec(ParseString(1)) + 1
                .Shiny = CBool(ParseString(1))
                .UnownLetter = Bin2Dec(ParseString(5))
                For Y = 1 To 4
                    .Move(Y) = Bin2Dec(ParseString(9))
                    Call SetPokePP(PNum, X, Y, PPUp)
                    .PP(Y) = Bin2Dec(ParseString(7))
                Next Y
                .DV_HP = Bin2Dec(ParseString(5))
                .DV_Atk = Bin2Dec(ParseString(5))
                .DV_Def = Bin2Dec(ParseString(5))
                .DV_Spd = Bin2Dec(ParseString(5))
                .DV_SAtk = Bin2Dec(ParseString(5))
                .DV_SDef = Bin2Dec(ParseString(5))
                .EV_HP = Bin2Dec(ParseString(8))
                .EV_Atk = Bin2Dec(ParseString(8))
                .EV_Def = Bin2Dec(ParseString(8))
                .EV_Spd = Bin2Dec(ParseString(8))
                .EV_SAtk = Bin2Dec(ParseString(8))
                .EV_SDef = Bin2Dec(ParseString(8))
                Call FillInPokeData(PKMN(PNum, X), .GameVersion)
                .Condition = Bin2Dec(ParseString(4))
                .ConditionCount = Bin2Dec(ParseString(4))
                .HP = Bin2Dec(ParseString(BinDig(.MaxHP)))
                .ItemEffect = Bin2Dec(ParseString(8))
                .RecycleItem = Bin2Dec(ParseString(7))
                .Resting = Bin2Bool(ParseString(1))
                Y = Bin2Dec(ParseString(4))
                .Nickname = Bin2Chr(ParseString(Y * 8))
                If .Nickname = "" Then .Nickname = .Name
                If UseBG Then Y = 3 Else Y = GameVersion(PNum)
                'Y = GameVersion(PNum)
                .Image = ChooseImage(PKMN(PNum, X), Y)
                .TeamNumber = X
            End With
        Next X
    Next PNum
    
    RealWeather = Bin2Dec(ParseString(3))
    WeatherCount = Bin2Dec(ParseString(3))
    TurnNumber = Bin2Dec(ParseString(10))
    Terrain = Bin2Dec(ParseString(4))
    BattlePauseNum = Bin2Dec(ParseString(3))
    For X = 1 To ActNum
        Order(X).MoveRandom = Bin2Dec(ParseString(3))
        Order(X).EndTurnRandom = Bin2Dec(ParseString(3))
        Order(X).PriorityTier = Bin2Dec(ParseString(4)) - 5
        Unfreezing(X) = Bin2Bool(ParseString(1))
        SwitchedThisTurn(X) = Bin2Bool(ParseString(1))
    Next X
    For X = 1 To 2
        BlankPos(X) = Bin2Dec(ParseString(3))
        SelfLoss(X) = Bin2Bool(ParseString(1))
    Next X
        
    For X = 1 To ActNum
        Current(X) = PKMN(TeamNum(X), Bin2Dec(ParseString(3)))
        Y = Bin2Dec(ParseString(4))
        BCondition(X).TransformedTo = Y
        If Y > 0 Then
            BuildPKMN = Current(X)
            If Y > 6 Then
                Current(X) = PKMN(TeamNum(X), Y - 6)
            Else
                Current(X) = PKMN(OtherTeam(TeamNum(X)), Y)
            End If
            With Current(X)
                .TeamNumber = BuildPKMN.TeamNumber
                .Nickname = BuildPKMN.Nickname
                .Name = BuildPKMN.Name
                .Item = BuildPKMN.Item
                .ItemEffect = BuildPKMN.ItemEffect
                .MaxHP = BuildPKMN.MaxHP
                .HP = BuildPKMN.HP
                .Level = BuildPKMN.Level
                .Gender = BuildPKMN.Gender
                .DV_HP = BuildPKMN.DV_HP
                .DV_Atk = BuildPKMN.DV_Atk
                .DV_Def = BuildPKMN.DV_Def
                .DV_Spd = BuildPKMN.DV_Spd
                .DV_SAtk = BuildPKMN.DV_SAtk
                .DV_SDef = BuildPKMN.DV_SDef
                .EV_HP = BuildPKMN.EV_HP
                .EV_Atk = BuildPKMN.EV_Atk
                .EV_Def = BuildPKMN.EV_Def
                .EV_Spd = BuildPKMN.EV_Spd
                .EV_SAtk = BuildPKMN.EV_SAtk
                .EV_SDef = BuildPKMN.EV_SDef
                .Condition = BuildPKMN.Condition
                .ConditionCount = BuildPKMN.ConditionCount
                .Resting = BuildPKMN.Resting
                .RecycleItem = BuildPKMN.RecycleItem
                .Image = ChooseImage(Current(X), GameVersion(TeamNum(X)))
                For Y = 1 To 4
                    .PP(Y) = BuildPKMN.PP(Y)
                Next Y
            End With
        End If
    Next X
        
    For X = 1 To ActNum
        Current(X).Type1 = Bin2Dec(ParseString(5))
        Current(X).Type2 = Bin2Dec(ParseString(5))
        Current(X).Attribute = Bin2Dec(ParseString(7))
        With BCondition(X)
            .AttackChange = Bin2Dec(ParseString(4)) - 6
            .DefenseChange = Bin2Dec(ParseString(4)) - 6
            .SpeedChange = Bin2Dec(ParseString(4)) - 6
            .SAttackChange = Bin2Dec(ParseString(4)) - 6
            .SDefenseChange = Bin2Dec(ParseString(4)) - 6
            .AccuracyChange = Bin2Dec(ParseString(4)) - 6
            .EvadeChange = Bin2Dec(ParseString(4)) - 6
            .StuckMove = Bin2Dec(ParseString(9))
            .StuckCount = Bin2Dec(ParseString(3))
            .RepeatMove = Bin2Dec(ParseString(9))
            .RepeatCount = Bin2Dec(ParseString(3))
            .DisabledMove = Bin2Dec(ParseString(3))
            .DisableCount = Bin2Dec(ParseString(4))
            .EncoreMove = Bin2Dec(ParseString(3))
            .EncoreDuration = Bin2Dec(ParseString(3))
            .BideDamage = Bin2Dec(ParseString(10))
            .BideCount = Bin2Dec(ParseString(2))
            .LockOnCount = Bin2Dec(ParseString(2))
            .ConfuseCounter = Bin2Dec(ParseString(4)) - 1
            .Attract = Expand(Bin2Dec(ParseString(2)), X)
            .LeechSeed = Expand(Bin2Dec(ParseString(2)), X)
            .TrapNum = Expand(Bin2Dec(ParseString(2)), X)
            .LockOn = Expand(Bin2Dec(ParseString(2)), X)
            .RepeatPos = Expand(Bin2Dec(ParseString(2)), X)
            .ToxicCount = Bin2Dec(ParseString(4))
            .PerishSong = Bin2Dec(ParseString(3))
            .Rollout = Bin2Dec(ParseString(3))
            .FuryCutter = Bin2Dec(ParseString(3))
            .ProtectPercent = Bin2Dec(ParseString(7))
            .RageCounter = Bin2Dec(ParseString(6))
            .MimicedMove = Bin2Dec(ParseString(9))
            .Substitute = Bin2Dec(ParseString(10))
            .LastMoveSlot = Bin2Dec(ParseString(3))
            .Stockpile = Bin2Dec(ParseString(2))
            .YawnCount = Bin2Dec(ParseString(2))
            .YawnSleepDuration = Bin2Dec(ParseString(4))
            .ChargeCount = Bin2Dec(ParseString(2))
            .WishCount = Bin2Dec(ParseString(2))
            .WishPos = Bin2Dec(ParseString(3))
            .SemiInvul = Bin2Dec(ParseString(3))
            .Snatch = Bin2Dec(ParseString(3))
            .TraceNumber = Bin2Dec(ParseString(1))
            .OutrageConfuse = Bin2Dec(ParseString(4))
            .TauntCount = Bin2Dec(ParseString(3))
            .ChoiceBandSlot = Bin2Dec(ParseString(3))
            .FakeOut = Bin2Dec(ParseString(2))
            .Confuse = Bin2Bool(ParseString(1))
            .Charging = Bin2Bool(ParseString(1))
            .Recharging = Bin2Bool(ParseString(1))
            .FocusEnergy = Bin2Bool(ParseString(1))
            .Foresight = Bin2Bool(ParseString(1))
            .DestinyBond = Bin2Bool(ParseString(1))
            .Nightmare = Bin2Bool(ParseString(1))
            .DefenseCurl = Bin2Bool(ParseString(1))
            .Curse = Bin2Bool(ParseString(1))
            .Minimize = Bin2Bool(ParseString(1))
            .Mist = Bin2Bool(ParseString(1))
            .Encore = Bin2Bool(ParseString(1))
            .BurnPenalty = Bin2Bool(ParseString(1))
            .ParalyzePenalty = Bin2Bool(ParseString(1))
            .BatonPassing = Bin2Bool(ParseString(1))
            .InPursuit = Bin2Bool(ParseString(1))
            .BeingPursued = Bin2Bool(ParseString(1))
            .Flinching = Bin2Bool(ParseString(1))
            .Protected = Bin2Bool(ParseString(1))
            .Enduring = Bin2Bool(ParseString(1))
            .Damaged = Bin2Bool(ParseString(1))
            .HelpingHand = Bin2Bool(ParseString(1))
            .Grudge = Bin2Bool(ParseString(1))
            .Ingrain = Bin2Bool(ParseString(1))
            .Truant = Bin2Bool(ParseString(1))
            .ShedSkin = Bin2Bool(ParseString(1))
            .MagicCoat = Bin2Bool(ParseString(1))
            .FlashFire = Bin2Bool(ParseString(1))
            .DoneCheck = Bin2Bool(ParseString(1))
            .SwitchedThisTurn = Bin2Bool(ParseString(1))
            .LansatBerry = Bin2Bool(ParseString(1))
            .WaterSport = Bin2Bool(ParseString(1))
            .MudSport = Bin2Bool(ParseString(1))
            .Torment = Bin2Bool(ParseString(1))
            .FaintLast = Bin2Bool(ParseString(1))
            .TraceOK = Bin2Bool(ParseString(1))
            .FollowMe = Bin2Bool(ParseString(1))
            .Imprison = Bin2Bool(ParseString(1))
            If .Imprison Then
                For Y = 1 To 4
                    .SealedMove(Y) = ParseString(9)
                Next Y
            End If
            If ParseString(1) = "1" Then
                .FutureSight.AccuracyMod = Bin2Dec(ParseString(4)) - 6
                .FutureSight.AttackPower = Bin2Dec(ParseString(13))
                .FutureSight.CHit = Bin2Bool(ParseString(1))
                .FutureSight.count = Bin2Dec(ParseString(3))
                .FutureSight.FSMove = Bin2Dec(ParseString(1))
                .FutureSight.Level = Bin2Dec(ParseString(7))
                .FutureSight.RandomNumber = Bin2Dec(ParseString(8))
            End If
        End With
    Next X
    For X = 1 To 2
        With TeamBCondition(X)
            .Spikes = Bin2Dec(ParseString(2))
            .ReflectCount = Bin2Dec(ParseString(3))
            .LightScreenCount = Bin2Dec(ParseString(3))
            .SafeguardCount = Bin2Dec(ParseString(3))
            .MistCount = Bin2Dec(ParseString(3))
        End With
    Next X
    
    For PNum = 1 To 2
        For X = NumPoke + 1 To 6
            PKMN(PNum, X) = BlankPKMN
            PKMN(PNum, X).Condition = 8
        Next X
    Next PNum
End Property
Public Property Get BattleSync() As String
    Dim Temp As String
    Dim Final As String
    Dim TArray() As Integer
    Dim BArray() As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim PNum As Byte
   
    For PNum = 1 To 2
                
        'PP (6 bits each * 6 Pokemon * 4 Moves = 144 bits)
        For X = 1 To 6
            For Y = 1 To 4
                Temp = Temp & Dec2Bin(PKMN(PNum, X).PP(Y), 6)
            Next Y
        Next X
        
        'HP (10 bits each * 6 Pokemon = 60 bits)
        For X = 1 To 6
            Temp = Temp & Dec2Bin(PKMN(PNum, X).HP, 10)
        Next X
        
        'Items (6 bits each * 6 Pokemon = 36 bits)
        For X = 1 To 6
            Temp = Temp & Dec2Bin(PKMN(PNum, X).Item, 6)
        Next X
        
        'Conditions and Counts (6 bits each * 6 Pokemon = 36 bits)
        For X = 1 To 6
            Y = PKMN(PNum, X).Condition
            If Y <> 0 Then Y = Y - 1
            Temp = Temp & Dec2Bin(Y, 3)
            Temp = Temp & Dec2Bin(PKMN(PNum, X).ConditionCount, 4)
        Next X
        
        'Current Pokemon (3 bits)
        Temp = Temp & Dec2Bin(Current(PNum).TeamNumber, 3)
        
        With BCondition(PNum)
            Temp = Temp & Dec2Bin(.AttackChange + 6, 4) '//4
            Temp = Temp & Dec2Bin(.DefenseChange + 6, 4) '//4
            Temp = Temp & Dec2Bin(.SpeedChange + 6, 4) '//4
            Temp = Temp & Dec2Bin(.SAttackChange + 6, 4) '//4
            Temp = Temp & Dec2Bin(.SDefenseChange + 6, 4) '//4
            Temp = Temp & Dec2Bin(.AccuracyChange + 6, 4) '//4
            Temp = Temp & Dec2Bin(.EvadeChange + 6, 4) '//4
            Temp = Temp & Dec2Bin(.MoveUsed(1), 9) '//9
            Temp = Temp & Dec2Bin(.MoveUsed(2), 9) '//9
            'Temp = Temp & Dec2Bin(.LastDamage, 10) '//10
            'Temp = Temp & Dec2Bin(.LastSDamage, 10) '//10
            Temp = Temp & Dec2Bin(.StuckMove, 9) '//9
            Temp = Temp & Dec2Bin(.StuckCount, 3) '//3
            Temp = Temp & Dec2Bin(.RepeatMove, 9) '//9
            Temp = Temp & Dec2Bin(.RepeatCount, 3) '//3
            Temp = Temp & Dec2Bin(.DisabledMove, 9) '//9
            Temp = Temp & Dec2Bin(.EncoreMove, 7)  '//7
            Temp = Temp & Dec2Bin(.EncoreDuration, 3)  '//3
            Temp = Temp & Dec2Bin(.DisableCount, 4) '//4
            Temp = Temp & Dec2Bin(.BideDamage, 10) '//10
            Temp = Temp & Dec2Bin(.BideCount, 2) '//2
            Temp = Temp & Dec2Bin(.ToxicCount, 5) '//5
            Temp = Temp & Dec2Bin(.PerishSong, 2) '//2
            Temp = Temp & Dec2Bin(.Rollout, 3) '//3
            Temp = Temp & Dec2Bin(.FuryCutter, 3) '//3
            Temp = Temp & Dec2Bin(.Substitute, 8) '//8
            Temp = Temp & Dec2Bin(.ProtectPercent, 7) '//7
            Temp = Temp & Dec2Bin(.MimicedMove, 9) '//9
            Temp = Temp & Dec2Bin(.RageCounter, 7) '//7
            Temp = Temp & Dec2Bin(.TransformedTo, 3) '//3
            Temp = Temp & Dec2Bin(.ConfuseCounter + 1, 4) '//4
'            Temp = Temp & Dec2Bin(.WrapCount, 3)   '//3
'            Temp = Temp & Dec2Bin(.WrapMove, 9)    '//9
            Temp = Temp & CStr(Bool2Bin(.Confuse)) '//1
            Temp = Temp & CStr(Bool2Bin(.LeechSeed)) '//1
            Temp = Temp & CStr(Bool2Bin(.Attract)) '//1
            Temp = Temp & CStr(Bool2Bin(.Charging)) '//1
            Temp = Temp & CStr(Bool2Bin(.Recharging)) '//1
'            Temp = Temp & CStr(Bool2Bin(.Dig)) '//1
'            Temp = Temp & CStr(Bool2Bin(.Fly)) '//1
'            Temp = Temp & CStr(Bool2Bin(.Locked))  '//1
            Temp = Temp & CStr(Bool2Bin(.FocusEnergy)) '//1
            Temp = Temp & CStr(Bool2Bin(.LockOn)) '//1
            Temp = Temp & CStr(Bool2Bin(.Foresight)) '//1
            Temp = Temp & CStr(Bool2Bin(.DestinyBond)) '//1
            Temp = Temp & CStr(Bool2Bin(.Curse)) '//1
            Temp = Temp & CStr(Bool2Bin(.Minimize)) '//1
            Temp = Temp & CStr(Bool2Bin(.Encore)) '//1
            Temp = Temp & CStr(Bool2Bin(.Nightmare)) '//1
            Temp = Temp & CStr(Bool2Bin(.DefenseCurl)) '//1
            Temp = Temp & CStr(Bool2Bin(.BurnPenalty)) '//1
            Temp = Temp & CStr(Bool2Bin(.ParalyzePenalty))  '//1
            'Temp = Temp & CStr(Bool2Bin(.Wrapped)) '//1
            Temp = Temp & CStr(Bool2Bin(.BatonPassing))  '//1
'            Temp = Temp & Dec2Bin(.FutureSight.Attacker, 3) '//3
'            Temp = Temp & Dec2Bin(.FutureSight.SAtkMod + 6, 4) '//4
'            Temp = Temp & Dec2Bin(.FutureSight.Turn, 10) '//10
'            Temp = Temp & CStr(Bool2Bin(.FutureSight.Hit)) '//1
'            Temp = Temp & CStr(Bool2Bin(.FutureSight.CHit)) '//1
        End With
        With TeamBCondition(PNum)
            Temp = Temp & CStr(Bool2Bin(.Spikes))  '//1
            Temp = Temp & Dec2Bin(.ReflectCount, 3)  '//3
            Temp = Temp & Dec2Bin(.LightScreenCount, 3)   '//3
            Temp = Temp & Dec2Bin(.SafeguardCount, 3)  '//3
        End With
    Next PNum
    Temp = Temp & Dec2Bin(TurnNumber, 10) '//10
    Temp = Temp & Dec2Bin(RealWeather, 2) '//2
    Temp = Temp & Dec2Bin(WeatherCount, 3) '//3
    
    'Debug.Print "BattleSync Binary: " & Temp
    'Debug.Print "Length: " & Len(Temp)
    Temp = Bin2Chr(Temp)
    BattleSync = Temp
End Property

Public Property Let BattleSync(SyncString As String)
    Dim Temp As String
    Dim Final As String
    Dim TArray() As Integer
    Dim BArray() As Boolean
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim PNum As Byte
    
    Temp = Left(Chr2Bin(SyncString), 1059)
    For PNum = 1 To 2
        
        'PP (6 bits each * 6 Pokemon * 4 Moves = 144 bits)
        For X = 1 To 6
            For Y = 1 To 4
                PKMN(PNum, X).PP(Y) = Bin2Dec(ChopString(Temp, 6))
            Next Y
        Next X
        
        'HP (10 bits each * 6 Pokemon = 60 bits)
        For X = 1 To 6
            PKMN(PNum, X).HP = Bin2Dec(ChopString(Temp, 10))
        Next X
        
        'Items (6 bits each * 6 Pokemon = 36 bits)
        For X = 1 To 6
            PKMN(PNum, X).Item = Bin2Dec(ChopString(Temp, 6))
        Next X
        
        'Conditions and Counts (6 bits each * 6 Pokemon = 36 bits)
        For X = 1 To 6
            PKMN(PNum, X).Condition = Bin2Dec(ChopString(Temp, 3)) + 1
            PKMN(PNum, X).ConditionCount = Bin2Dec(ChopString(Temp, 4))
        Next X
        
        'Current Pokemon (3 bits)
        Current(PNum) = PKMN(PNum, Bin2Dec(ChopString(Temp, 3)))

        With BCondition(PNum)
            .AttackChange = Bin2Dec(ChopString(Temp, 4)) - 6
            .DefenseChange = Bin2Dec(ChopString(Temp, 4)) - 6
            .SpeedChange = Bin2Dec(ChopString(Temp, 4)) - 6
            .SAttackChange = Bin2Dec(ChopString(Temp, 4)) - 6
            .SDefenseChange = Bin2Dec(ChopString(Temp, 4)) - 6
            .AccuracyChange = Bin2Dec(ChopString(Temp, 4)) - 6
            .EvadeChange = Bin2Dec(ChopString(Temp, 4)) - 6
            .MoveUsed(1) = Bin2Dec(ChopString(Temp, 9))
            .MoveUsed(2) = Bin2Dec(ChopString(Temp, 9))
            '.LastDamage = Bin2Dec(ChopString(Temp, 10))
            '.LastSDamage = Bin2Dec(ChopString(Temp, 10))
            .StuckMove = Bin2Dec(ChopString(Temp, 9))
            .StuckCount = Bin2Dec(ChopString(Temp, 3))
            .RepeatMove = Bin2Dec(ChopString(Temp, 9))
            .RepeatCount = Bin2Dec(ChopString(Temp, 3))
            .DisabledMove = Bin2Dec(ChopString(Temp, 9))
            .EncoreMove = Bin2Dec(ChopString(Temp, 7))
            .EncoreDuration = Bin2Dec(ChopString(Temp, 3))
            .DisableCount = Bin2Dec(ChopString(Temp, 4))
            .BideDamage = Bin2Dec(ChopString(Temp, 10))
            .BideCount = Bin2Dec(ChopString(Temp, 2))
            .ToxicCount = Bin2Dec(ChopString(Temp, 5))
            .PerishSong = Bin2Dec(ChopString(Temp, 2))
            .Rollout = Bin2Dec(ChopString(Temp, 3))
            .FuryCutter = Bin2Dec(ChopString(Temp, 3))
            .Substitute = Bin2Dec(ChopString(Temp, 8))
            .ProtectPercent = Bin2Dec(ChopString(Temp, 7))
            .MimicedMove = Bin2Dec(ChopString(Temp, 9))
            .RageCounter = Bin2Dec(ChopString(Temp, 7))
            .TransformedTo = Bin2Dec(ChopString(Temp, 3))
            .ConfuseCounter = Bin2Dec(ChopString(Temp, 4)) - 1
'            .WrapCount = Bin2Dec(ChopString(Temp, 3))
'            .WrapMove = Bin2Dec(ChopString(Temp, 9))
            .Confuse = CBool(ChopString(Temp, 1))
            .LeechSeed = CBool(ChopString(Temp, 1))
            .Attract = CBool(ChopString(Temp, 1))
            .Charging = CBool(ChopString(Temp, 1))
            .Recharging = CBool(ChopString(Temp, 1))
'            .Dig = CBool(ChopString(Temp, 1))
'            .Fly = CBool(ChopString(Temp, 1))
'            .Locked = CBool(ChopString(Temp, 1))
            .FocusEnergy = CBool(ChopString(Temp, 1))
            .LockOn = CBool(ChopString(Temp, 1))
            .Foresight = CBool(ChopString(Temp, 1))
            .DestinyBond = CBool(ChopString(Temp, 1))
            .Curse = CBool(ChopString(Temp, 1))
            .Minimize = CBool(ChopString(Temp, 1))
            .Encore = CBool(ChopString(Temp, 1))
            .Nightmare = CBool(ChopString(Temp, 1))
            .DefenseCurl = CBool(ChopString(Temp, 1))
            .BurnPenalty = CBool(ChopString(Temp, 1))
            .ParalyzePenalty = CBool(ChopString(Temp, 1))
            '.Wrapped = CBool(ChopString(Temp, 1))
            .BatonPassing = CBool(ChopString(Temp, 1))
'            .FutureSight.Attacker = Bin2Dec(ChopString(Temp, 3))
'            .FutureSight.SAtkMod = Bin2Dec(ChopString(Temp, 4)) - 6
'            .FutureSight.Turn = Bin2Dec(ChopString(Temp, 10))
'            .FutureSight.Hit = CBool(ChopString(Temp, 1))
'            .FutureSight.CHit = CBool(ChopString(Temp, 1))
        End With
        With TeamBCondition(PNum)
            .Spikes = CBool(ChopString(Temp, 1))
            .ReflectCount = Bin2Dec(ChopString(Temp, 3))
            .LightScreenCount = Bin2Dec(ChopString(Temp, 3))
            .SafeguardCount = Bin2Dec(ChopString(Temp, 3))
        End With
    Next PNum
        
    For PNum = 1 To 2
        If BCondition(PNum).TransformedTo > 0 Then
            Y = Current(PNum).TeamNumber
            Z = Current(PNum).ConditionCount
            Current(PNum) = PKMN(OtherTeam(PNum), BCondition(PNum).TransformedTo)
            With Current(PNum)
                .TeamNumber = Y
                .Nickname = PKMN(PNum, Y).Nickname
                .Item = PKMN(PNum, Y).Item
                .MaxHP = PKMN(PNum, Y).MaxHP
                .HP = PKMN(PNum, Y).HP
                .DV_HP = PKMN(PNum, Y).DV_HP
                .DV_Atk = PKMN(PNum, Y).DV_Atk
                .DV_Def = PKMN(PNum, Y).DV_Def
                .DV_Spd = PKMN(PNum, Y).DV_Spd
                .DV_SAtk = PKMN(PNum, Y).DV_SAtk
                .Condition = PKMN(PNum, Y).Condition
                .ConditionCount = Y
                If UseBG Then X = 3 Else X = GameVersion(PNum)
                .Image = ChooseImage(Current(PNum), X)
            End With
        End If
    Next PNum
    
    TurnNumber = Bin2Dec(ChopString(Temp, 10))
    RealWeather = Bin2Dec(ChopString(Temp, 2))
    WeatherCount = Bin2Dec(ChopString(Temp, 3))
    If InVBMode And Temp <> "" Then Stop
    'Debug.Print "BattleSync OK!"
End Property

Sub SetPokePP(ByVal Team As Byte, ByVal Position As Byte, ByVal MoveNum As Byte, ByVal UsePPUps As Boolean)
    Dim X As Single
    
    With PKMN(Team, Position)
        If UsePPUps Then X = 1.6 Else X = 1
        .PP(MoveNum) = Int(Moves(.Move(MoveNum)).PP * X)
        If .PP(MoveNum) = 64 And BattleMode <> nbAdvBattle Then .PP(MoveNum) = 61
        .MaxPP(MoveNum) = .PP(MoveNum)
    End With
End Sub

Public Sub AddMessage(ByVal TextIndex As Integer, _
Optional Param1Type As AddMessageTypes, Optional ByVal Param1Num As Long, _
Optional Param2Type As AddMessageTypes, Optional ByVal Param2Num As Long, _
Optional Param3Type As AddMessageTypes, Optional ByVal Param3Num As Long, _
Optional Param4Type As AddMessageTypes, Optional ByVal Param4Num As Long, _
Optional DebugMessage As Boolean)
    Dim FlavorText As String
    Dim Param1 As String
    Dim Param2 As String
    Dim Param3 As String
    Dim Param4 As String
    Dim Bold As Boolean
    Dim Italic As Boolean
    Dim Color As Long
    Dim BreakChar As String
    Dim Temp As String
    Dim TempType As AddMessageTypes
    Dim X As Long
    Dim p As Long
    Dim Y As Long
    Dim Z As Long
    Dim F As Long
    If Len(LogFile) > 0 Then
        On Error GoTo Locked
        F = FreeFile
        Open LogFile For Binary Access Write As #F
        p = LOF(F)
        If DontTouchMe <> 8 And DontTouchMe <> 0 Then
            TextLog = Right$(TextLog, 1)
        Else
            p = p + 1
            TextLog = vbNullString
        End If
        SetDestStringASM VarPtr(TextLog), DontTouchMe
        StreamInASM TextIndex, 9
        X = 0
        If Param3Type = nbBlank Then
            X = 1
        ElseIf Param4Type = nbBlank Then
            X = 2
        Else
            X = 3
        End If
        StreamInASM X, 2
        For X = 0 To X
            Select Case X
            Case 0: TempType = Param1Type: Y = Param1Num
            Case 1: TempType = Param2Type: Y = Param2Num
            Case 2: TempType = Param3Type: Y = Param3Num
            Case 3: TempType = Param4Type: Y = Param4Num
            End Select
            StreamInASM TempType, 4
            Select Case TempType
            Case nbActive
                If TeamNum(Y) = 2 Then Z = 6 Else Z = 0
                Z = Z + Current(Y).TeamNumber
                StreamInASM Z, 4
            Case nbMove, nbPoke: StreamInASM Y, 9
            Case nbPlayer: StreamInASM Minimum(Y - 1), 1
            Case nbCond: StreamInASM Minimum(Y - 1), 3
            Case nbItem, nbTrait: StreamInASM Y, 7
            Case nbStat, nbRule: StreamInASM Y, 4
            Case nbNumber: StreamInASM Y, 10
            Case nbType: StreamInASM Y, 5
            End Select
        Next X
        Put #F, 1, CStr(DontTouchMe)
        Put #F, p, TextLog
        Close #F
    End If
Locked:

    If Not DoMessages Then Exit Sub
    If DebugMessage And Not DebugMode Then Exit Sub
    FlavorText = Replace(GetFText(TextIndex), "%n", vbNewLine)
    For X = 4 To 1 Step -1
        Select Case X
        Case 1: TempType = Param1Type: Y = Param1Num
        Case 2: TempType = Param2Type: Y = Param2Num
        Case 3: TempType = Param3Type: Y = Param3Num
        Case 4: TempType = Param4Type: Y = Param4Num

        End Select
        Select Case TempType
            Case nbActive
                If UseNicks Or TextIndex = 298 Or TextIndex = 285 Then
                    Temp = Current(Y).Nickname
                Else
                    Temp = Current(Y).Name
                End If
                Select Case TextIndex
                Case 25 To 28, 285, 298
                Case Else
                    If UsePrefix Then
                        Temp = PName(TeamNum(Y)) & "'s " & Temp
                    End If
                End Select
            Case nbMove: Temp = Moves(Y).Name
            Case nbPlayer: Temp = PName(Y)
            Case nbCond: Temp = Condition(Y)
            Case nbItem: Temp = Item(Y)
            Case nbStat: Temp = StatName(Y)
            Case nbNumber: Temp = CStr(Y)
            Case nbPoke: Temp = BasePKMN(Y).Name
            Case nbRule: Temp = RuleText(Y)
            Case nbTrait: Temp = AttributeText(Y)
            Case nbType: Temp = Element(Y)
            Case Else: Temp = ""
        End Select
        FlavorText = Replace(FlavorText, "%" & CStr(X), Temp, , 1)
    Next X
    
    Select Case TextIndex
    Case 5: Color = vbRed 'Critical Hit
    Case 23, 312 To 314: Color = vbMagenta: Bold = True: BreakChar = ":" 'Rule:
    Case 4, 24 To 36: Bold = True 'General stuff
    End Select
    
    ''Debug.Print FlavorText
    If Not DoMessages Then Exit Sub
    Call LogOutput.AddMessage(FlavorText, BreakChar, Color, Bold, Italic)
    
End Sub

Sub DoPPChange(ByVal Team As Byte, ByVal Slot As Integer, Optional ByVal PPChange As Integer = -1)
    Dim X As Integer
    Dim Y As Integer
    With Current(Team)
        X = TeamNum(Team)
        Y = .TeamNumber
        If BattleMode = nbRBYBattle And .PP(Slot) = 0 And PPChange = -1 Then
            .PP(Slot) = 63
            .MaxPP(Slot) = Moves(.Move(Slot)).PP * 1.6
            PKMN(X, Y).MaxPP(Slot) = Moves(PKMN(X, Y).Move(Slot)).PP * 1.6
        ElseIf .PP(Slot) + PPChange < 0 Then
            .PP(Slot) = 0
        Else
            .PP(Slot) = .PP(Slot) + PPChange
        End If
        If BCondition(Slot).TransformedTo = 0 Then
            If PKMN(X, Y).Move(Slot) <> .Move(Slot) Then
                If BattleMode = nbRBYBattle And PKMN(X, Y).Move(Slot) = 124 Then
                    PKMN(X, Y).PP(Slot) = Current(Team).PP(Slot)
                End If
            Else
                PKMN(X, Y).PP(Slot) = Current(Team).PP(Slot)
            End If
        End If
    End With
End Sub


Public Function DoThreePKMN() As Boolean
    Dim Team As Byte
    Dim SwapPKMN(3) As Pokemon
    Dim BlankPKMN As Pokemon
    
    For Team = 1 To 2
        SwapPKMN(1) = PKMN(Team, StoreStadium(Team, 1))
        SwapPKMN(2) = PKMN(Team, StoreStadium(Team, 2))
        SwapPKMN(3) = PKMN(Team, StoreStadium(Team, 3))
        PKMN(Team, 1) = SwapPKMN(1)
        PKMN(Team, 1).TeamNumber = 1
        PKMN(Team, 2) = SwapPKMN(2)
        PKMN(Team, 2).TeamNumber = 2
        PKMN(Team, 3) = SwapPKMN(3)
        PKMN(Team, 3).TeamNumber = 3
        PKMN(Team, 4) = BlankPKMN
        PKMN(Team, 4).Condition = 8
        PKMN(Team, 5) = BlankPKMN
        PKMN(Team, 5).Condition = 8
        PKMN(Team, 6) = BlankPKMN
        PKMN(Team, 6).Condition = 8
        Current(Team) = PKMN(Team, 1)
    Next
    DoThreePKMN = True
End Function

Private Sub Class_Initialize()
    NumPoke = 6
    TurnNumber = 1
    Set LogOutput = Nothing
    Set BattleWindow = Nothing
End Sub

Function CompleteBString()
    Dim X As Integer
    X = InStr(1, BattleString, ";")
    If InStr(X + 1, BattleString, ";") = 0 Then CompleteBString = False Else CompleteBString = True
End Function

Sub StartBattle()
    Dim Temp() As Byte
    Dim X As Long
    Call AddMessage(33, nbPlayer, 1, nbPlayer, 2)
    Call DoStartSwitch(1)
    Call DoStartSwitch(2)
    Temp = GetRandomOrder
    For X = 1 To ActNum
        Order(X).EndTurnRandom = Temp(X)
    Next X
    Call SyncPKMN
    If Not (BattleWindow Is Nothing) Then
        Call BattleWindow.UpdateImages
        Call BattleWindow.UpdateStats
    End If
    If UseDX Then
'        For X = 1 To 6
'            BattleWindow.DX.Surface(X).Visible = False
'        Next X
    End If
'    If ActNum = 4 Then
'        Call DoAnim(BattleWindow, nbaSwitch + 2, 1, Int(Rnd * 12), Int(Rnd * 12))
'        Call DoAnim(BattleWindow, nbaSwitch + 2, 2, Int(Rnd * 12), Int(Rnd * 12))
'    Else
'        Call DoAnim(BattleWindow, nbaSwitch + 1, 1, Int(Rnd * 12), Int(Rnd * 12))
'        Call DoAnim(BattleWindow, nbaSwitch + 1, 2, Int(Rnd * 12), Int(Rnd * 12))
'    End If
'    Call DoSwitch(1, 1, False, True)
'    Call DoSwitch(2, 1, False, True)
End Sub

Function ResetBattle()
    Dim BlankPKMN As Pokemon
    Dim BlankCondition As BattleStuff
    Dim BlankTC As TeamCond
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As Boolean
    Timeout = 0
    WaitingFor = 0
    TextLog = vbNullString
    DisconStall = False
    WatchIgnore1 = False
    WatchIgnore2 = False
    BattlePauseNum = 0
    BattleString = ""
    StoredWinner = 0
    DoMessages = False
    DoReplay = False
    Unrated = False
    Terrain = 0
    WeatherCount = 0
    GoesFirstOK = False
    GoesLastOK = False
    BattleMode = nbGSCBattle
    NumPoke = 6
    StadiumMode = False
    PPUp = False
    TurnNumber = 1
    Player1 = 0
    Player2 = 0
    RealWeather = nbNoWeather
    LogFile = vbNullString
    WatchOK = False
    ReDim PendingWatch(0)
    ReDim NowWatching(0)
    For X = 1 To UBound(RuleSelected)
        RuleSelected(X) = False
    Next
    For X = 1 To 2
        For Y = 1 To 6
            PKMN(X, Y) = BlankPKMN
        Next
        For Y = 1 To 3
            StoreStadium(X, Y) = 0
        Next
        TeamBCondition(X) = BlankTC
        CheckedIn(X) = False
'        BatonPass(X) = False
        PName(X) = ""
        SwitchedThisTurn(X) = False
        GotTeam(X) = False
        GameVersion(X) = 0
        NeedSelect(X) = False
        BlankPos(X) = 0
    Next
    For X = 1 To 4
        Moved(X) = False
        SwitchTo(X) = 0
        BCondition(X) = BlankCondition
        Current(X) = BlankPKMN
    Next X
    Err.Number = 0
    On Error Resume Next
    Temp = LogOutput.RTBSet
    DoMessages = (Err.Number = 0)
End Function
Private Sub ClearStuckMove(Team As Byte)
    'In BattleMode = nbRBYBattle Rage is not cleared under any circumstances.
    If BattleMode = nbRBYBattle And BCondition(Team).StuckMove = 156 Then Exit Sub
    If BCondition(Team).StuckMove > 0 Then BCondition(Team).StuckMove = 0
    If BCondition(Team).BideCount > 0 Then BCondition(Team).BideCount = 0
    If BattleMode = nbRBYBattle Then
        If BCondition(Team).RepeatMove > 0 Then BCondition(Team).RepeatMove = 0: BCondition(Team).RepeatCount = 0
    End If
End Sub
Public Sub NewWatcher(ByVal PNum As Long)
    ReDim Preserve NowWatching(UBound(NowWatching) + 1)
    NowWatching(UBound(NowWatching)) = PNum
End Sub
Public Property Let IsCancellable(ByVal Pos As Long, ByVal C As Boolean)
    Cancellable(Pos) = C
End Property
Public Property Get IsCancellable(ByVal Pos As Long) As Boolean
    IsCancellable = Cancellable(Pos)
End Property

Public Sub RemoveWatcher(ByVal PNum As Long)
    Dim X As Integer
    On Error GoTo None
    For X = 1 To UBound(NowWatching)
        If NowWatching(X) = PNum Then Exit For
    Next X
    If X <= UBound(NowWatching) Then
        For X = X To UBound(NowWatching) - 1
            NowWatching(X) = NowWatching(X + 1)
        Next X
        ReDim Preserve NowWatching(UBound(NowWatching) - 1)
    End If
None:
End Sub
Public Function IsWatching(ByVal PNum As Long) As Boolean
    Dim X As Integer
    On Error GoTo None
    For X = 1 To UBound(NowWatching)
        If NowWatching(X) = PNum Then
            IsWatching = True
            Exit Function
        End If
    Next X
    IsWatching = False
None:
End Function
Public Function GetWatchers() As String
    Dim T() As String
    Dim X As Integer
    If UBound(NowWatching) = 0 Then
        GetWatchers = ""
        Exit Function
    End If
    ReDim T(UBound(NowWatching) - 1)
    For X = 1 To UBound(NowWatching)
        T(X - 1) = CStr(NowWatching(X))
    Next X
    GetWatchers = Join(T, ";")
End Function
Public Function GetTT(ByVal PNum As Long, TeamPos As Byte, Optional HideStats As Boolean = False) As String
    Dim TempLine() As String
    Dim TStat(6) As String
    Dim CondLine() As String
    Dim Build As String
    Dim LBPika As Boolean
    Dim TCWak As Boolean
    Dim MPDitto As Boolean
    Dim TempStat As Integer
    Dim X As Integer
    Dim Y As Integer
    Const Prefix = "Under the effects of "
    Const Suffix = " is in effect"
    On Error GoTo ETrap
    'The spacing of the ToolTips is optimized for the default
    'font for ToolTips in WinXP, Tahoma.
    If TeamPos = 0 Then
        With Current(PNum)
            TCWak = ((.No = 104 Or .No = 105) And .Item = nbThickClub)
            LBPika = (.No = 25 And .Item = nbLightBall)
            MPDitto = (PKMN(TeamNum(PNum), Current(PNum).TeamNumber).No = 132 And .Item = nbMetalPowder)
        End With
        With BCondition(PNum)
            If HideStats Then
                TStat(0) = IIf(.AttackChange >= 0, "+", "-") & .AttackChange
                TStat(1) = IIf(.DefenseChange >= 0, "+", "-") & .DefenseChange
                TStat(2) = IIf(.SpeedChange >= 0, "+", "-") & .SpeedChange
                TStat(3) = IIf(.SAttackChange >= 0, "+", "-") & .SAttackChange
                TStat(4) = IIf(.SDefenseChange >= 0, "+", "-") & .SDefenseChange
            Else
                TStat(0) = GetAtk(PNum)
                TStat(1) = GetDef(PNum)
                TStat(2) = GetSpd(PNum)
                TStat(3) = GetSAtk(PNum)
                TStat(4) = GetSDef(PNum)
                Y = 3
                For X = 0 To 4
                    If Len(TStat(X)) > Y Then Y = Len(TStat(X))
                Next X
                For X = 0 To 4
                    TStat(X) = String$((Y - Len(TStat(X))) * 2, " ") & TStat(X)
                Next X
                
                If .AttackChange <> 0 Then TStat(0) = TStat(0) & " (" & IIf(.AttackChange >= 0, "+", "-") & .AttackChange & ")"
                If .DefenseChange <> 0 Then TStat(1) = TStat(1) & " (" & IIf(.DefenseChange >= 0, "+", "-") & .DefenseChange & ")"
                If .SpeedChange <> 0 Then TStat(2) = TStat(2) & " (" & IIf(.SpeedChange >= 0, "+", "-") & .SpeedChange & ")"
                If .SAttackChange <> 0 Then TStat(3) = TStat(3) & " (" & IIf(.SAttackChange >= 0, "+", "-") & .SAttackChange & ")"
                If .SDefenseChange <> 0 Then TStat(4) = TStat(4) & " (" & IIf(.SDefenseChange >= 0, "+", "-") & .SDefenseChange & ")"
            End If

            'Yes, the double - is on purpose.
            TStat(5) = IIf(Y = 4, "  ", "") & IIf(.AccuracyChange >= 0, "+", "-") & .AccuracyChange
            TStat(6) = IIf(Y = 4, "  ", "") & IIf(.EvadeChange >= 0, "+", "-") & .EvadeChange
        End With
        ReDim TempLine(IIf(BattleMode = nbRBYBattle, 5, 6))
        TempLine(0) = "Attack    ....... " & TStat(0)
        TempLine(1) = "Defense ....... " & TStat(1)
        TempLine(2) = "Speed   ........ " & TStat(2)
        If BattleMode = nbRBYBattle Then
            TempLine(3) = "Special  ........ " & TStat(3)
            TempLine(4) = "Accuracy ......" & IIf(HideStats, " ", ". ") & TStat(5)
            TempLine(5) = "Evasion  ......." & IIf(HideStats, " ", ". ") & TStat(6)
        Else
            TempLine(3) = "S.Attack  ...... " & TStat(3)
            TempLine(4) = "S.Defense   ... " & TStat(4)
            TempLine(5) = "Accuracy ......" & IIf(HideStats, " ", ". ") & TStat(5)
            TempLine(6) = "Evasion  ......." & IIf(HideStats, " ", ". ") & TStat(6)
        End If
        X = 0
        With BCondition(PNum)
            If .Confuse Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(29).Name '"Confusion"
            End If
            If .Attract <> 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(8).Name '"Attract"
            End If
            If .Curse Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(38).Name '"Curse"
            End If
            If Current(PNum).Condition = 3 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(235).Name '"Toxic"
            End If
            If .Nightmare Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(134).Name '"Nightmare"
            End If
            If .LeechSeed <> 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(106).Name '"Leech Seed"
            End If
            If .RepeatMove > 0 And BattleMode <> nbRBYBattle Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(.RepeatMove).Name
            End If
            If BattleMode = nbRBYBattle Then
                If BCondition(OtherTeam(PNum)).RepeatMove > 0 Then
                    X = X + 1: ReDim Preserve CondLine(X)
                    CondLine(X) = Prefix & Moves(BCondition(OtherTeam(PNum)).RepeatMove).Name
                End If
            End If
            If .StuckMove > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = "Stuck using " & Moves(.StuckMove).Name
            End If
'            If .WrapMove > 0 Then
'                X = X + 1: ReDim Preserve CondLine(X)
'                CondLine(X) = "Stuck using " & Moves(.WrapMove).Name
'            End If
            If .EncoreMove > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(58).Name '"Encore"
            End If
            If .TrapNum > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & "a Full Trapping Move"
            End If
            If .LockOn > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(110).Name & "/" & Moves(125).Name  '"Lock On/Mindreader"
            End If
            If .DestinyBond Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(41).Name '"Destiny Bond"
            End If
            If .Mist Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(129).Name '"Mist"
            End If
            If .FocusEnergy Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(73).Name '"Focus Energy"
            End If
            If .Foresight Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(74).Name '"Foresight"
            End If
            If BattleMode <> nbRBYBattle Then
                If .DefenseCurl Then
                    X = X + 1: ReDim Preserve CondLine(X)
                    CondLine(X) = Prefix & Moves(40).Name '"Defense Curl"
                End If
                If .Minimize Then
                    X = X + 1: ReDim Preserve CondLine(X)
                    CondLine(X) = Prefix & Moves(126).Name '"Minimize"
                End If
            End If
            If .HelpingHand Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(290).Name '"Helping Hand"
            End If
            If .Stockpile > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = "Stockpile level is " & .Stockpile
            End If
            If .ChargeCount > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(267).Name '"Charge"
            End If
            If .Ingrain Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(297).Name '"Ingrain"
            End If
            If .FlashFire Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & AttributeText(nbFlashFire)
            End If
            If .LansatBerry Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Item(nbLansatBerry)
            End If
            If .Imprison Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(296).Name '"Imprison"
            End If
            If .Torment Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(344).Name '"Torment"
            End If
            If .TauntCount > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Prefix & Moves(341).Name '"Taunt"
            End If
            If .DisabledMove > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Moves(Current(PNum).Move(.DisabledMove)).Name & " is Disabled"
            End If
            If TeamBCondition(TeamNum(PNum)).ReflectCount <> 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Moves(162).Name & Suffix '"Reflect"
            End If
            If TeamBCondition(TeamNum(PNum)).LightScreenCount <> 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Moves(109).Name & Suffix '"Light Screen"
            End If
            If TeamBCondition(TeamNum(PNum)).MistCount <> 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Moves(129).Name & Suffix '"Mist"
            End If
            If TeamBCondition(TeamNum(PNum)).SafeguardCount <> 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Moves(173).Name & Suffix '"Safeguard"
            End If
            If TeamBCondition(TeamNum(PNum)).Spikes <> 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = Moves(202).Name & Suffix '"Spikes"
                If BattleMode = nbAdvBattle Then
                    CondLine(X) = CondLine(X) & " (Level " & TeamBCondition(TeamNum(PNum)).Spikes & ")"
                End If
            End If
'            If .Substitute > 0 Then
'                X = X + 1: ReDim Preserve CondLine(X)
'                CondLine(X) = "Substitute health is " & CStr(.Substitute) & "/" & CStr(Current(PNum).MaxHP \ 4) & " HP"
'            End If
            If .PerishSong > 0 Then
                X = X + 1: ReDim Preserve CondLine(X)
                CondLine(X) = "Perish Count is " & CStr(.PerishSong)
            End If
        End With
        On Error Resume Next
        CondLine(0) = vbNewLine & vbNewLine & "Current Conditions:"
        Build = Join(TempLine, vbNewLine)
        Build = Build & Join(CondLine, vbNewLine)
        GetTT = Build
    Else
        With PKMN(PNum, TeamPos)
            Build = "Lv." & CStr(.Level) & " " & PKMN(PNum, TeamPos).Name & vbNewLine
            Build = Build & "Type: " & Element(.Type1)
            If .Type2 > nbNoType Then Build = Build & "/" & Element(.Type2)
            If BattleMode = nbAdvBattle Then
                Build = Build & vbNewLine & "Trait: " & AttributeText(.Attribute)
                'Build = Build & vbNewLine & "EVs: " & .EV_HP & "/" & .EV_Atk & "/" & .EV_Def & "/" & .EV_Spd & "/" & .EV_SAtk & "/" & .EV_SDef
            End If
            Build = Build & vbNewLine
            For X = 1 To 4
                If .Move(X) > 0 Then
                    Build = Build & vbNewLine & Moves(.Move(X)).Name & " - " & .PP(X) & " PP"
                End If
            Next X
            If BattleMode <> nbRBYBattle Then
                Build = Build & vbNewLine & "[" & IIf(.Item = nbNoItem, "No Item", Item(.Item)) & "]"
            End If
        End With
        GetTT = Build
    End If
    Exit Function
ETrap:
    If InVBMode Then
        Stop
        Resume
    End If
End Function
Public Function AllyNum(PokeNum As Byte) As Byte
    If PokeNum > 2 Then AllyNum = PokeNum - 2 Else AllyNum = PokeNum + 2
End Function
Public Function GetItemEffect() As Byte
    Dim X As Integer
    Dim M(1 To 4) As Integer
    Dim Build As String
    M(1) = 1
    M(2) = 2
    M(3) = 2
    M(4) = 3
    For X = 4 To 1 Step -1
        Build = Build & Dec2Bin(GetSmallRnd(CByte(X)), M(X))
    Next X
    GetItemEffect = Bin2Dec(Build)
End Function
Public Function ItemEffect2Array(iEffect As Byte) As Byte()
    Dim Build(1 To 5) As Byte
    Dim Offset(1 To 5) As Byte
    Dim Taken(1 To 5) As Boolean
    Dim Temp As String
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Temp = Dec2Bin(iEffect, 8)
    Offset(1) = Bin2Dec(ChopString(Temp, 3))
    Offset(2) = Bin2Dec(ChopString(Temp, 2))
    Offset(3) = Bin2Dec(ChopString(Temp, 2))
    Offset(4) = Bin2Dec(ChopString(Temp, 1))
    For X = 1 To 5
        Z = Offset(X) + 1
        For Y = 1 To 5
            If Not Taken(Y) Then Z = Z - 1
            If Z = 0 Then Exit For
        Next Y
        Build(X) = Y
        Taken(Y) = True
    Next X
    ItemEffect2Array = Build
End Function
Public Function GetAtk(ByVal Pos As Byte, Optional UseMods As Boolean = True) As Integer
    Dim X As Integer
    Dim Y As Integer
    With Current(Pos)
        X = .Attack
        If UseMods Then X = X * StatChange(BCondition(Pos).AttackChange)
        Y = PKMN(TeamNum(Pos), .TeamNumber).No
        If (Y = 104 Or Y = 105) And .Item = nbThickClub Then X = Rollover(X * 2)
        If .Attribute = nbHugePower Then X = X * 2
        If .Attribute = nbPurePower Then X = X * 2
        If .Attribute = nbGuts And .Condition > 1 And .Condition < 8 Then X = Int(X * 1.5)
        If .Item = nbChoiceBand Then X = X * 1.5
        If BattleMode <> nbAdvBattle Then X = Cap(X)
        If BCondition(Pos).BurnPenalty And .Attribute <> nbGuts Then X = X \ 2
    End With
    GetAtk = X
End Function
Public Function GetDef(ByVal Pos As Byte, Optional UseMods As Boolean = True) As Integer
    Dim X As Integer
    Dim Y As Integer
    With Current(Pos)
        X = .Defense
        If UseMods Then X = X * StatChange(BCondition(Pos).DefenseChange)
        Y = PKMN(TeamNum(Pos), .TeamNumber).No
        If Y = 132 And .Item = nbMetalPowder Then X = Rollover(X * 1.5)
        If .Attribute = nbMarvelScale And .Condition > 1 And .Condition < 8 Then X = Int(X * 1.5)
        If BattleMode <> nbAdvBattle Then X = Cap(X)
    End With
    GetDef = X
End Function
Public Function GetSAtk(ByVal Pos As Byte, Optional UseMods As Boolean = True) As Integer
    Dim X As Integer
    Dim Y As Integer
    With Current(Pos)
        X = .SpecialAttack
        If UseMods Then X = X * StatChange(BCondition(Pos).SAttackChange)
        Y = PKMN(TeamNum(Pos), .TeamNumber).No
        If Y = 25 And .Item = nbLightBall Then X = Rollover(X * 2)
        If Y = 366 And .Item = nbDeepSeaTooth Then X = X * 2
        If (Y = 380 Or Y = 381) And .Item = nbSoulDew Then X = X * 1.5
        If ActNum = 4 Then
            If Current(AllyNum(Pos)).HP > 0 Then
                If (Current(Pos).Attribute = nbPlus And Current(AllyNum(Pos)).Attribute = nbMinus) _
                Or (Current(Pos).Attribute = nbMinus And Current(AllyNum(Pos)).Attribute = nbPlus) Then
                    X = X * 1.5
                End If
            End If
        End If
        If BattleMode <> nbAdvBattle Then X = Cap(X)
    End With
    GetSAtk = X
End Function
Public Function GetSDef(ByVal Pos As Byte, Optional UseMods As Boolean = True) As Integer
    Dim X As Integer
    Dim Y As Integer
    With Current(Pos)
        X = .SpecialDefense
        If UseMods Then X = X * StatChange(BCondition(Pos).SDefenseChange)
        Y = PKMN(TeamNum(Pos), .TeamNumber).No
        If Y = 132 And .Item = nbMetalPowder Then X = Rollover(X * 1.5)
        If Y = 366 And .Item = nbDeepSeaScale Then X = X * 2
        If (Y = 380 Or Y = 381) And .Item = nbSoulDew Then X = X * 1.5
        If BattleMode <> nbAdvBattle Then X = Cap(X)
    End With
    GetSDef = X
End Function
Public Function GetSpd(ByVal Pos As Byte, Optional UseMods As Boolean = True) As Integer
    Dim X As Integer
    With Current(Pos)
        X = .Speed
        If UseMods Then X = X * StatChange(BCondition(Pos).SpeedChange)
        If Current(Pos).Attribute = nbSwiftSwim And CurrentWeather = nbRaining Then X = X * 2
        If Current(Pos).Attribute = nbChlorophyll And CurrentWeather = nbSunny Then X = X * 2
        If BattleMode <> nbAdvBattle Then X = Cap(X)
        If BCondition(Pos).ParalyzePenalty Then X = X \ 4
        If Current(Pos).Item = nbMachoBrace Then X = X \ 2
    End With
    GetSpd = X
End Function
Function AttackBase(ByVal MoveType As Elements) As AttackBases
    Select Case MoveType
    Case 2 To 6, 11, 15, 16
        AttackBase = nbSpecialBased
    Case Else
        AttackBase = nbPysicalBased
    End Select
End Function


Sub DoEndRoundGSC(ResumeNum As Byte)
    Dim X As Byte
    If ResumeNum < 2 Then
        For X = 1 To ActNum
            If Unfreezing(X) And Current(X).Condition = 7 Then
                Call AddMessage(185, nbActive, X)
                Current(X).Condition = 1
            End If
            Unfreezing(X) = False
        Next X
        Call DoFutureSight
        ResumeNum = ResumeNum + 1: If NeedSwitch Then Exit Sub
    End If
    If ResumeNum < 3 Then
        Call DoWeather
        ResumeNum = ResumeNum + 1: If NeedSwitch Then Exit Sub
    End If
    If ResumeNum < 4 Then
        Call DoRepeatMoves
        ResumeNum = ResumeNum + 1: If NeedSwitch Then Exit Sub
    End If
    If ResumeNum < 5 Then
        If BattleMode = nbGSCBattle Then Call DoBarrierFade
        Call DoEncore
        Call DoNightmare
        ResumeNum = ResumeNum + 1: If NeedSwitch Then Exit Sub
    End If
    If ResumeNum < 6 Then
        Call DoPerishSong
        ResumeNum = ResumeNum + 1: If NeedSwitch Then Exit Sub
    End If
    If ResumeNum < 7 Then
        Call DoLeftovers
        ResumeNum = ResumeNum + 1
        Call DoEndRoundCommon
    End If
End Sub
Sub DoEndRoundAdv(ResumeNum As Byte)
    Dim X As Byte
    If ResumeNum < 2 Then
        Call DoBarrierFade
        Call DoWish
        Call DoWeather
        ResumeNum = ResumeNum + 1: If NeedSwitch Then Exit Sub
    End If
    If ResumeNum < 3 Then
        Call DoConditionsAdv
        ResumeNum = ResumeNum + 1: If NeedSwitch Then Exit Sub
    End If
    If ResumeNum < 4 Then
        Call DoFutureSight
        If NeedSwitch Then Exit Sub Else ResumeNum = ResumeNum + 1
    End If
    If ResumeNum < 5 Then
        Call DoYawn
        ResumeNum = ResumeNum + 1
    End If
    If ResumeNum < 6 Then
        Call DoPerishSong
        If NeedSwitch Then Exit Sub Else ResumeNum = ResumeNum + 1
    End If
    If ResumeNum < 7 Then
        Call DoEncore
        For X = 1 To ActNum
            With BCondition(X)
                If .ChargeCount > 0 Then .ChargeCount = .ChargeCount - 1
                If .TauntCount > 0 Then .TauntCount = .TauntCount - 1
                If .LockOnCount > 0 Then
                    .LockOnCount = .LockOnCount - 1
                    If .LockOnCount = 0 Then .LockOn = 0
                End If
                .HelpingHand = False
            End With
        Next X
        ResumeNum = ResumeNum + 1
        Call DoEndRoundCommon
    End If
End Sub
Private Sub DoEndRoundCommon()
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Dim A As Long
    If BMessStyle = 1 Then
        Call AddMessage(24, nbNumber, TurnNumber)
        For Y = 1 To 2
            For X = Y To ActNum Step 2
                A = Current(X).HP
                If A > 0 Then
                    If Y <> ThisPNum And Not EnabledRule(nbExactHP) Then
                        A = Round(A * 100 / Current(X).MaxHP)
                        If Current(X).HP > 0 And A = 0 Then A = 1
                        Z = 1
                    Else
                        Z = 0
                    End If
                    If Current(X).Condition = 1 Then
                        Call AddMessage(25 + Z, nbPlayer, TeamNum(X), nbActive, X, nbNumber, A)
                    Else
                        Call AddMessage(27 + Z, nbPlayer, TeamNum(X), nbActive, X, nbNumber, A, nbCond, Current(X).Condition)
                    End If
                End If
            Next X
        Next Y
    End If
    TurnNumber = TurnNumber + 1
    For X = 1 To ActNum
        If Not (Ready(X) Or BCondition(X).Encore) Then BCondition(X).TargetNum = 0
    Next X
End Sub
'End Round Subs
'-----------------------------------------------------
Sub DoConditionsGSC(AtkPos As Byte)
    'Everything that happens after each Pokemon's turn happens here.
    'This is the GSC and RBY version.
    Dim Temp As Integer
    Dim B As Boolean
    Dim X As Byte
    Dim Y As Integer
    Dim Z As Integer
    Dim T As Byte
    'First, check for Items
    Call DoItems
    
    Z = IIf(BattleMode = nbRBYBattle, 16, 8)
    T = BCondition(AtkPos).ToxicCount
    B = (BattleMode = nbRBYBattle And T > 0)
    Select Case Current(AtkPos).Condition
    'Poison/Burn Damage
    Case 2, 5
        If Not B Then T = 1 Else BCondition(AtkPos).ToxicCount = BCondition(AtkPos).ToxicCount + 1: T = T + 1
        Y = Int(T * Current(AtkPos).MaxHP / Z)
        If Y = 0 Then Y = 1
        Current(AtkPos).HP = Current(AtkPos).HP - Y
        Call AddMessage(IIf(Current(AtkPos).Condition = 2, 176, 181), nbActive, AtkPos)
    'Toxic Damage
    Case 3
        BCondition(AtkPos).ToxicCount = BCondition(AtkPos).ToxicCount + 1
        Y = BCondition(AtkPos).ToxicCount * Int(Current(AtkPos).MaxHP / 16)
        If Y = 0 Then Y = 1
        Current(AtkPos).HP = Current(AtkPos).HP - Y
        Call AddMessage(176, nbActive, AtkPos)
'    'Burn Damage
'    Case 5
'        Current(AtkPos).HP = Current(AtkPos).HP - Current(AtkPos).MaxHP / 8
'        Call AddMessage(181, nbActive, AtkPos)
    End Select
    Call SyncPKMN: If NeedSwitch Then Exit Sub
    
    T = BCondition(AtkPos).ToxicCount
    B = (BattleMode = nbRBYBattle And T > 0)
    'Leech Seed Damage
    If BCondition(AtkPos).LeechSeed <> 0 Then
        If Not B Then T = 1 Else BCondition(AtkPos).ToxicCount = BCondition(AtkPos).ToxicCount + 1: T = T + 1
        Y = Int(T * Current(AtkPos).MaxHP / Z)
        If Y = 0 Then Y = 1
        X = BCondition(AtkPos).LeechSeed
        Current(X).HP = Current(X).HP + Cap(Y, Current(AtkPos).HP)
        Current(AtkPos).HP = Current(AtkPos).HP - Y
        Call AddMessage(211, nbActive, AtkPos, nbMove, 106)
    End If
    Call SyncPKMN: If NeedSwitch Then Exit Sub
    
    'Curse Damage
    If BCondition(AtkPos).Curse Then
        Call ChangeHP(AtkPos, -4)
        Call AddMessage(201, nbActive, AtkPos)
    End If
    Call SyncPKMN
End Sub
Sub ChangeHP(ByVal Pos As Byte, ByVal Fraction As Long)
    Dim X As Long
    X = Current(Pos).MaxHP \ Fraction
    If X = 0 Then X = Sgn(Fraction)
    Current(Pos).HP = Current(Pos).HP + X
    If Current(Pos).HP > Current(Pos).MaxHP Then Current(Pos).HP = Current(Pos).MaxHP
End Sub
Sub DoConditionsAdv()
    Dim X As Byte
    Dim i As Byte
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Byte
    Dim O() As Byte
    O = GetOrderArray
    For A = 1 To ActNum
        X = O(A)
        
        If Current(X).HP > 0 Then
            If Current(X).Attribute = nbRainDish And CurrentWeather = nbRaining And Current(X).HP < Current(X).MaxHP Then
                Call AddMessage(156, nbActive, X, nbTrait, nbRainDish)
                Call ChangeHP(X, 16)
                Call SyncPKMN
            ElseIf Current(X).Attribute = nbSpeedBoost And BCondition(X).FakeOut <> 2 Then
                Call AddMessage(133, nbActive, X, nbTrait, nbSpeedBoost)
                Call ModifyPKMN(X, 7, 1, True)
            ElseIf Current(X).Attribute = nbShedSkin And BCondition(X).ShedSkin And Current(X).Condition <> 1 Then
                Select Case Current(X).Condition
                Case 2, 3: Call AddMessage(159, nbActive, X, nbTrait, nbShedSkin)
                Case 4: Call AddMessage(299, nbActive, X, nbTrait, nbShedSkin)
                Case 5: Call AddMessage(167, nbActive, X, nbTrait, nbShedSkin)
                Case 6: Call AddMessage(300, nbActive, X, nbTrait, nbShedSkin)
                Case 7: Call AddMessage(301, nbActive, X, nbTrait, nbShedSkin)
                End Select
                Current(X).Condition = 1
                Current(X).ConditionCount = 0
                Current(X).Resting = False
                BCondition(X).BurnPenalty = False
                BCondition(X).ParalyzePenalty = False
                Call SyncPKMN
            ElseIf Current(X).Attribute = nbTruant Then
                BCondition(X).Truant = Not BCondition(X).Truant
            Else
                BCondition(X).Truant = False
            End If
            BCondition(X).ShedSkin = False
            
            'Ingrain
            If BCondition(X).Ingrain And Current(X).HP < Current(X).MaxHP Then
                Call AddMessage(89, nbActive, X)
                Call ChangeHP(X, 16)
                Call SyncPKMN
            End If
    
            Call DoPinchItems(X)
            
            'Leftovers
            If Current(X).Item = nbLeftovers And Current(X).HP < Current(X).MaxHP Then
                Call AddMessage(130, nbActive, X, nbItem, 21)
                Call ChangeHP(X, 16)
                Call SyncPKMN
            End If
    
            'Leech Seed Damage
            If BCondition(X).LeechSeed <> 0 Then
                If Current(BCondition(X).LeechSeed).HP > 0 Then
                    Z = Int(Current(X).MaxHP / 8)
                    If Z = 0 Then Z = 1
                    Y = BCondition(X).LeechSeed
                    Call AddMessage(211, nbActive, X, nbMove, 106)
                    If Current(X).Attribute = nbLiquidOoze Then
                        Call AddMessage(160, nbActive, BCondition(X).LeechSeed, nbTrait, nbLiquidOoze)
                        Current(Y).HP = Current(Y).HP - Cap(Z, Current(X).HP)
                    Else
                        Current(Y).HP = Current(Y).HP + Cap(Z, Current(X).HP)
                    End If
                    Current(X).HP = Current(X).HP - Z
                    Call SyncPKMN: If NeedSwitch Then GoTo EndLoop
                End If
            End If
            
            Select Case Current(X).Condition
            'Poison Damage
            Case 2
                Call ChangeHP(X, -8)
                Call AddMessage(176, nbActive, X)
                Call SyncPKMN: If NeedSwitch Then GoTo EndLoop
            'Toxic Damage
            Case 3
                BCondition(X).ToxicCount = BCondition(X).ToxicCount + 1
                Y = BCondition(X).ToxicCount * (Current(X).MaxHP / 16)
                If Y = 0 Then Y = 1
                Current(X).HP = Current(X).HP - Y
                Call AddMessage(176, nbActive, X)
                Call SyncPKMN: If NeedSwitch Then GoTo EndLoop
            'Burn Damage
            Case 5
                Call ChangeHP(X, -8)
                Call AddMessage(181, nbActive, X)
                Call SyncPKMN: If NeedSwitch Then GoTo EndLoop
            End Select
    
            'Nightmare Damage
            If BCondition(X).Nightmare Then
                If Current(X).Condition = 4 Then
                    Call ChangeHP(X, -4)
                    Call AddMessage(199, nbActive, X)
                    Call SyncPKMN: If NeedSwitch Then GoTo EndLoop
                Else
                    BCondition(X).Nightmare = False
                End If
            End If
    
            'Curse Damage
            If BCondition(X).Curse Then
                Call ChangeHP(X, -4)
                Call AddMessage(201, nbActive, X)
                Call SyncPKMN: If NeedSwitch Then GoTo EndLoop
            End If
            
            'Partial Trapping Move Damage
            If BCondition(X).RepeatMove > 0 Then
                BCondition(X).RepeatCount = BCondition(X).RepeatCount - 1
                If BCondition(X).RepeatCount = 0 Then
                    Call AddMessage(208, nbActive, X, nbMove, BCondition(X).RepeatMove)
                    BCondition(X).RepeatMove = 0
                    BCondition(X).RepeatPos = 0
                Else
                    Call AddMessage(207, nbActive, X, nbMove, BCondition(X).RepeatMove)
                    Call ChangeHP(X, -16)
                    Call SyncPKMN: If NeedSwitch Then GoTo EndLoop
                End If
            End If
            
            'StuckMoves
            If BCondition(X).StuckCount > 0 Then
                BCondition(X).StuckCount = BCondition(X).StuckCount - 1
                Select Case Moves(BCondition(X).StuckMove).SpecialEffect
                Case 79 'Outrage/Thrash/Petal Dance
                    If BCondition(X).StuckCount = 0 Then
                        BCondition(X).StuckMove = 0
                        If Not BCondition(X).Confuse And TeamBCondition(TeamNum(X)).SafeguardCount = 0 And Current(X).Attribute <> nbOwnTempo Then
                            Call AddMessage(195, nbActive, X)
                            BCondition(X).Confuse = True
                            BCondition(X).ConfuseCounter = BCondition(X).OutrageConfuse
                        End If
                    End If
                Case 174 'Uproar
                    If BCondition(X).StuckCount = 0 Then
                        BCondition(X).StuckMove = 0
                        Call AddMessage(226, nbActive, X)
                    Else
                        For Y = 1 To ActNum
                            If Current(Y).Condition = 4 And Current(Y).Attribute <> nbCacophony And Current(Y).Attribute <> nbSoundproof Then
                                Current(Y).Condition = 1
                                Call AddMessage(233, nbActive, Y)
                            End If
                        Next Y
                        Call AddMessage(225, nbActive, X)
                    End If
                End Select
            End If
            
            'Disable
            If BCondition(X).DisabledMove > 0 Then
               BCondition(X).DisableCount = BCondition(X).DisableCount - 1
                If BCondition(X).DisableCount = 0 Then
                    Call AddMessage(235, nbActive, X)
                    BCondition(X).DisabledMove = 0
                End If
            End If
        End If
EndLoop:
    Next A
End Sub
Sub DoWish()
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Dim O() As Byte
    Dim T1 As String
    Dim T2 As String
    Dim T3 As Byte
    O = GetOrderArray
    For Y = 1 To ActNum
        X = O(Y)
        If BCondition(X).WishCount <> 0 Then
            BCondition(X).WishCount = BCondition(X).WishCount - 1
            If BCondition(X).WishCount = 0 And Current(X).HP > 0 Then
                Z = BCondition(X).WishPos
                With Current(X)
                    'See Beat Up's comment for why this is here
                    T1 = .Nickname: T2 = .Name: T3 = .TeamNumber
                    .Nickname = PKMN(TeamNum(X), Z).Nickname
                    .Name = PKMN(TeamNum(X), Z).Name
                    .TeamNumber = Z
                    Call AddMessage(87, nbActive, X)
                    .Nickname = T1: .Name = T2: .TeamNumber = T3
                End With
                If Current(X).HP < Current(X).MaxHP Then
                    Call AddMessage(80, nbActive, X)
                    Call ChangeHP(X, 2)
                    Call SyncPKMN
                Else
                    Call AddMessage(81, nbActive, X)
                End If
            End If
        End If
    Next Y
End Sub
Sub DoYawn()
    Dim X As Byte
    Dim Y As Byte
    Dim O() As Byte
    O = GetOrderArray
    For Y = 1 To ActNum
        X = O(Y)
        If BCondition(X).YawnCount <> 0 Then
            BCondition(X).YawnCount = BCondition(X).YawnCount - 1
            If BCondition(X).YawnCount = 0 Then
                If Current(X).Condition = 1 And UproarCheck(X) = 0 And Current(X).Attribute <> nbInsomnia And Current(X).Attribute <> nbVitalSpirit And Not RuleCheck(TeamNum(X), 1) Then
                    Call AddMessage(169, nbActive, X)
                    Current(X).Condition = 4
                    Current(X).ConditionCount = BCondition(X).YawnSleepDuration
                    Call SyncPKMN
                End If
                BCondition(X).YawnSleepDuration = 0
            End If
        End If
    Next Y
End Sub
Sub DoItems()
    Dim X As Byte
    Dim Y As Integer
    Dim i As Items
    Dim T As Items
    For X = 1 To ActNum
        If Current(X).HP > 0 Then
            i = Current(X).Item
            T = i
            Select Case i
            'GSC Healing Berries
            Case nbBerry, nbBerryJuice, nbGoldBerry
                If (Current(X).HP < Current(X).MaxHP / 2 Or Current(X).HP < Current(X).MaxHP - 128) And BattleMode <> nbAdvBattle Then
                    Call AddMessage(127, nbActive, X, nbItem, i)
                    Select Case i
                        Case nbBerry: Y = 10
                        Case nbBerryJuice: Y = 20
                        Case nbGoldBerry: Y = 30
                    End Select
                    Current(X).HP = Current(X).HP + Y
                    Current(X).Item = nbNoItem
                End If
            'Miracle Berry
            Case nbMiracleBerry, nbLumBerry
                If Current(X).Condition > 1 Or BCondition(X).Confuse Then
                    If BCondition(X).Confuse Then
                        Y = 302
                    Else
                        Select Case Current(X).Condition
                        Case 2, 3: Y = 159
                        Case 4: Y = 299
                        Case 5: Y = 167
                        Case 6: Y = 300
                        Case 7: Y = 301
                        End Select
                    End If
                    Call AddMessage(Y, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    Current(X).Condition = 1
                    Current(X).ConditionCount = 0
                    Current(X).Resting = False
                    BCondition(X).ParalyzePenalty = False
                    BCondition(X).BurnPenalty = False
                    BCondition(X).Confuse = False
                    BCondition(X).ConfuseCounter = 0
                End If
            'Poison Cure Berries
            Case nbPoisoncureBerry, nbPechaBerry
                If Current(X).Condition = 2 Or Current(X).Condition = 3 Then
                    Call AddMessage(120, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    Current(X).Condition = 1
                End If
            'Sleep Cure Berries
            Case nbMintBerry, nbChestoBerry
                If Current(X).Condition = 4 Then
                    Call AddMessage(123, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    Current(X).Condition = 1
                    Current(X).Resting = False
                    BCondition(X).Nightmare = False
                End If
            'Burn Cure Berries
            Case nbIceBerry, nbRawstBerry
                If Current(X).Condition = 5 Then
                    Call AddMessage(121, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    Current(X).Condition = 1
                    BCondition(X).BurnPenalty = False
                End If
            'Paralyze Cure Berries
            Case nbParalyzecureBerry, nbCheriBerry
                If Current(X).Condition = 6 Then
                    Call AddMessage(119, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    Current(X).Condition = 1
                    BCondition(X).ParalyzePenalty = False
                End If
            'Freeze Cure Berry
            Case nbBurntBerry, nbAspearBerry
                If Current(X).Condition = 7 Then
                    Call AddMessage(122, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    Current(X).Condition = 1
                End If
            'Confusion Cure Berry
            Case nbBitterBerry, nbPersimBerry
                If BCondition(X).Confuse Then
                    Call AddMessage(124, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    BCondition(X).Confuse = False
                    BCondition(X).ConfuseCounter = 0
                End If
            'PP Restore Berry (GSC only)
            Case nbMysteryBerry
                With PKMN(TeamNum(X), Current(X).TeamNumber)
                    For Y = 1 To 4
                        If .Move(Y) > 0 And .PP(Y) = 0 Then
                            Call AddMessage(128, nbActive, X, nbItem, i, nbMove, .Move(Y))
                            Current(X).Item = nbNoItem
                            .PP(Y) = 5
                            If .Move(Y) = Current(X).Move(Y) And BCondition(X).TransformedTo = 0 Then
                                Current(X).PP(Y) = Cap(5, Current(X).MaxPP(Y))
                            End If
                            Exit For
                        End If
                    Next Y
                End With
            'Attraction Cure
            Case nbMentalHerb
                If BCondition(X).Attract > 0 Then
                    Call AddMessage(303, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                    BCondition(X).Attract = 0
                End If
            'Stat Reduction Cure
            Case nbWhiteHerb
                Y = 0
                If BCondition(X).AttackChange < 0 Then BCondition(X).AttackChange = 0: Y = 1
                If BCondition(X).DefenseChange < 0 Then BCondition(X).DefenseChange = 0: Y = 1
                If BCondition(X).SpeedChange < 0 Then BCondition(X).SpeedChange = 0: Y = 1
                If BCondition(X).SAttackChange < 0 Then BCondition(X).SAttackChange = 0: Y = 1
                If BCondition(X).SDefenseChange < 0 Then BCondition(X).SDefenseChange = 0: Y = 1
                If BCondition(X).AccuracyChange < 0 Then BCondition(X).AccuracyChange = 0: Y = 1
                If BCondition(X).EvadeChange < 0 Then BCondition(X).EvadeChange = 0: Y = 1
                If Y = 1 Then
                    Call AddMessage(129, nbActive, X, nbItem, i)
                    Current(X).Item = nbNoItem
                End If
            End Select
            If T > nbNoItem And Current(X).Item = nbNoItem Then Current(X).RecycleItem = T
        End If
    Next X
End Sub
Sub DoPinchItems(Pos As Byte)
    Dim X As Integer
    Dim Y As Integer
    Dim i As Items
    Dim T As Items
    Dim E() As Byte
    'Ruby/Sapp Items that activate "in a pinch"
    i = Current(Pos).Item
    T = i
    E = ItemEffect2Array(Current(Pos).ItemEffect)
    If Current(Pos).HP <= Current(Pos).MaxHP \ 2 Then
        Select Case i
        Case nbOranBerry
            Call AddMessage(127, nbActive, Pos, nbItem, i)
            Current(Pos).HP = Current(Pos).HP + 10
            Current(Pos).Item = nbNoItem
        Case nbSitrusBerry
            Call AddMessage(127, nbActive, Pos, nbItem, i)
            Current(Pos).HP = Current(Pos).HP + 30
            Current(Pos).Item = nbNoItem
        Case nbFigyBerry, nbWikiBerry, nbMagoBerry, nbAguavBerry, nbIapapaBerry
            Call AddMessage(127, nbActive, Pos, nbItem, i)
            Call ChangeHP(Pos, 8)
            X = E(1)
            If X = 5 Then X = E(2)
            Y = CInt(i) - nbFigyBerry + 1
            If Nature(Current(Pos).NatureNum).StatChg(Y) = -1 Then Y = Y + 305 Else Y = 0
            If Y > 0 Then
                Call AddMessage(Y, nbActive, Pos, nbItem, i)
                If Not (BCondition(Pos).Confuse Or Current(Pos).Attribute = nbOwnTempo Or TeamBCondition(TeamNum(Pos)).SafeguardCount > 0) Then
                    Call AddMessage(193, nbActive, Pos)
                    BCondition(Pos).Confuse = True
                    BCondition(Pos).ConfuseCounter = X
                End If
            End If
            Current(Pos).Item = nbNoItem
        End Select
    End If
    If Current(Pos).HP <= Current(Pos).MaxHP \ 4 Then
        Select Case i
        Case nbLiechiBerry
            If BCondition(Pos).AttackChange < 6 Then
                Call AddMessage(304, nbItem, i, nbStat, 2, nbActive, Pos)
                Call ModifyPKMN(Pos, 1, 1, True)
                Current(Pos).Item = nbNoItem
            End If
        Case nbGanlonBerry
            If BCondition(Pos).DefenseChange < 6 Then
                Call AddMessage(304, nbItem, i, nbStat, 3, nbActive, Pos)
                Call ModifyPKMN(Pos, 2, 1, True)
                Current(Pos).Item = nbNoItem
            End If
        Case nbPetayaBerry
            If BCondition(Pos).SAttackChange < 6 Then
                Call AddMessage(304, nbItem, i, nbStat, 6, nbActive, Pos)
                Call ModifyPKMN(Pos, 3, 1, True)
                Current(Pos).Item = nbNoItem
            End If
        Case nbApicotBerry
            If BCondition(Pos).SDefenseChange < 6 Then
                Call AddMessage(304, nbItem, i, nbStat, 7, nbActive, Pos)
                Call ModifyPKMN(Pos, 4, 1, True)
                Current(Pos).Item = nbNoItem
            End If
        Case nbSalacBerry
            If BCondition(Pos).SpeedChange < 6 Then
                Call AddMessage(304, nbItem, i, nbStat, 4, nbActive, Pos)
                Call ModifyPKMN(Pos, 7, 1, True)
                Current(Pos).Item = nbNoItem
            End If
        Case nbLansatBerry
            If Not BCondition(Pos).LansatBerry Then
                Call AddMessage(283, nbActive, Pos, nbItem, i)
                BCondition(Pos).LansatBerry = True
                Current(Pos).Item = nbNoItem
            End If
        Case nbStarfBerry
            For X = 1 To 5
                Select Case E(X)
                Case 1
                    If BCondition(Pos).AttackChange < 6 Then
                        Call AddMessage(305, nbItem, i, nbStat, 2, nbActive, Pos)
                        Call ModifyPKMN(Pos, 1, 2, True)
                        Current(Pos).Item = nbNoItem
                        Exit For
                    End If
                Case 2
                    If BCondition(Pos).DefenseChange < 6 Then
                        Call AddMessage(305, nbItem, i, nbStat, 3, nbActive, Pos)
                        Call ModifyPKMN(Pos, 2, 2, True)
                        Current(Pos).Item = nbNoItem
                        Exit For
                    End If
                Case 3
                    If BCondition(Pos).SAttackChange < 6 Then
                        Call AddMessage(305, nbItem, i, nbStat, 6, nbActive, Pos)
                        Call ModifyPKMN(Pos, 3, 2, True)
                        Current(Pos).Item = nbNoItem
                        Exit For
                    End If
                Case 4
                    If BCondition(Pos).SDefenseChange < 6 Then
                        Call AddMessage(305, nbItem, i, nbStat, 7, nbActive, Pos)
                        Call ModifyPKMN(Pos, 4, 2, True)
                        Current(Pos).Item = nbNoItem
                        Exit For
                    End If
                Case 5
                    If BCondition(Pos).SpeedChange < 6 Then
                        Call AddMessage(305, nbItem, i, nbStat, 4, nbActive, Pos)
                        Call ModifyPKMN(Pos, 7, 2, True)
                        Current(Pos).Item = nbNoItem
                        Exit For
                    End If
                End Select
            Next X
        End Select
    End If
    If i = nbLeppaBerry Then
        With PKMN(TeamNum(Pos), Current(Pos).TeamNumber)
            For Y = 1 To 4
                If .Move(Y) > 0 And .PP(Y) = 0 Then
                    Call AddMessage(128, nbActive, Pos, nbItem, i, nbMove, .Move(Y))
                    Current(Pos).Item = nbNoItem
                    .PP(Y) = Cap(10, .MaxPP(Y))
                    If .Move(Y) = Current(Pos).Move(Y) And BCondition(Pos).TransformedTo = 0 Then
                        Current(Pos).PP(Y) = Cap(10, Current(Pos).MaxPP(Y))
                    End If
                    Exit For
                End If
            Next Y
        End With
    End If
    If T > nbNoItem And Current(Pos).Item = nbNoItem Then Current(Pos).RecycleItem = T
    Call SyncPKMN
End Sub
Sub DoFutureSight()
    Dim d As Byte
    Dim CHit As Boolean
    Dim ATK As Integer
    Dim POW As Integer
    Dim DEF As Integer
    Dim Lev As Integer
    Dim Damage As Integer
    Dim Hit As Boolean
    Dim X As Integer
    Dim Y As Byte
    Dim Z As Integer
    Dim O() As Byte
    Dim M As Integer
    Dim Ratio As Single
    O = GetOrderArray
    For Y = 1 To ActNum
        d = O(Y)
        If Not BCondition(d).DoneCheck Then
            BCondition(d).DoneCheck = True
            With BCondition(d).FutureSight
                If .count > 0 Then
                    .count = .count - 1
                    If .count = 0 And Current(d).HP > 0 Then
                        M = IIf(.FSMove = 0, 79, 272)
                        CHit = .CHit
                        ATK = .AttackPower
                        POW = Moves(M).Power
                        If AttackBase(Moves(M).Type) = nbPysicalBased Then
                            DEF = GetDef(d)
                        Else
                            DEF = GetSDef(d)
                        End If
                        Lev = .Level
                        Ratio = AccuracyMod(-BCondition(d).EvadeChange)
                        If Current(d).Item = nbBrightPowder Or Current(d).Item = nbLaxIncense Then Ratio = Ratio * 0.922
                        Hit = Odds(Moves(M).Accuracy * Ratio, CInt(.RandomNumber))
                        If CHit Then Lev = Lev + Lev
                        Call AddMessage(256, nbActive, d, nbMove, M)
                        Damage = ((((2 * Lev / 5 + 2) * ATK * POW) / DEF) / 50) + 2
                        If Hit = True And BCondition(d).SemiInvul = 0 And Damage > 0 Then
                            If BCondition(d).Substitute > 0 Then
                                BCondition(d).Substitute = BCondition(d).Substitute - Damage
                                Call AddMessage(100, nbActive, d)
                                If BCondition(d).Substitute <= 0 Then
                                    BCondition(d).Substitute = 0
                                    Call AddMessage(101, nbActive, d)
                                End If
                            Else
                                Current(d).HP = Current(d).HP - Damage
                                If Current(d).HP < 0 Then Damage = Damage + Current(d).HP
                                If BCondition(d).BideCount > 0 And BattleMode <> nbAdvBattle Then BCondition(d).BideDamage = BCondition(d).BideDamage + Damage
                                Call DoDamageMessage(d, Damage)
                                If CHit Then Call AddMessage(5)
                            End If
                        Else
                            Call AddMessage(38)
                        End If
                        Call SyncPKMN
                        If BattleMode = nbAdvBattle And NeedSwitch Then
                            If FaintedPokes(d) <> NumPoke Then Exit Sub
                        End If
                    End If
                End If
            End With
        End If
    Next Y
    For X = 1 To ActNum
        BCondition(X).DoneCheck = False
    Next X
End Sub
Sub DoWeather()
    'Count down the weather changes
    'Do damage with Sandstorm and Hail
    Dim X As Byte
    Dim Y As Byte
    Dim O() As Byte
    If WeatherCount = 7 Then WeatherCount = 8
    If WeatherCount > 0 Then WeatherCount = WeatherCount - 1
    If WeatherCount = 0 Then
        Select Case RealWeather
        Case 1: Call AddMessage(107)
        Case 2: Call AddMessage(114)
        Case 3: Call AddMessage(110)
        Case 4: Call AddMessage(117)
        End Select
        RealWeather = nbNoWeather
    Else
        Select Case RealWeather
        Case 1: Call AddMessage(105)
        Case 2: Call AddMessage(113)
        Case 3: Call AddMessage(109)
        Case 4: Call AddMessage(116)
        End Select
    End If
    
    Call DoForecast
    
    Select Case CurrentWeather
    Case nbSandstorm
        O = GetOrderArray
        For Y = 1 To ActNum
            X = O(Y)
            If Not TypeCheck(Current(X), nbRock, nbGround, nbSteel) _
            And Current(X).Attribute <> nbSandVeil And Current(X).HP > 0 Then
                Call AddMessage(111, nbActive, X)
                Call ChangeHP(X, IIf(BattleMode = nbAdvBattle, -16, -8))
                Call SyncPKMN
            End If
        Next Y
    Case nbHailstorm
        O = GetOrderArray
        For Y = 1 To ActNum
            X = O(Y)
            If Not TypeCheck(Current(X), nbIce) And Current(X).HP > 0 Then
                Call AddMessage(118, nbActive, X)
                Call ChangeHP(X, -16)
                Call SyncPKMN
            End If
        Next Y
    End Select
End Sub
Sub DoRepeatMoves()
    'Moves that hit at the end of a round
    '(Fire Spin, Wrap, Whirlpool, etc)
    Dim X As Byte
    Dim Y As Byte
    Dim O() As Byte
    If BattleMode = nbRBYBattle Then Exit Sub
    O = GetOrderArray
    For Y = 1 To ActNum
        X = O(Y)
        If BCondition(X).RepeatMove > 0 Then
            BCondition(X).RepeatCount = BCondition(X).RepeatCount - 1
            If BCondition(X).RepeatCount = 0 Then
                Call AddMessage(208, nbActive, X, nbMove, BCondition(X).RepeatMove)
                BCondition(X).RepeatMove = 0
                BCondition(X).RepeatPos = 0
            Else
                Call AddMessage(207, nbActive, X, nbMove, BCondition(X).RepeatMove)
                Current(X).HP = Current(X).HP - (Current(X).MaxHP / 16)
                Call SyncPKMN
            End If
        End If
    Next Y
End Sub
Sub DoBarrierFade()
    'Count down a barrier
    Dim X As Byte
    If BattleMode = nbRBYBattle Then Exit Sub
    
    'Reflect
    For X = 1 To 2
        If TeamBCondition(X).ReflectCount > 0 Then
            TeamBCondition(X).ReflectCount = TeamBCondition(X).ReflectCount - 1
            If TeamBCondition(X).ReflectCount = 0 Then Call AddMessage(12, nbPlayer, X, nbMove, 162)
        End If
    Next X
    
    'Light Screen
    For X = 1 To 2
        If TeamBCondition(X).LightScreenCount > 0 Then
            TeamBCondition(X).LightScreenCount = TeamBCondition(X).LightScreenCount - 1
            If TeamBCondition(X).LightScreenCount = 0 Then Call AddMessage(12, nbPlayer, X, nbMove, 109)
        End If
    Next X
    
    'Mist (Ru/Sa)
    For X = 1 To 2
        If TeamBCondition(X).MistCount > 0 Then
            TeamBCondition(X).MistCount = TeamBCondition(X).MistCount - 1
            If TeamBCondition(X).MistCount = 0 Then Call AddMessage(12, nbPlayer, X, nbMove, 129)
        End If
    Next X
    
    'Safegaurd
    For X = 1 To 2
        If TeamBCondition(X).SafeguardCount > 0 Then
            TeamBCondition(X).SafeguardCount = TeamBCondition(X).SafeguardCount - 1
            If TeamBCondition(X).SafeguardCount = 0 Then Call AddMessage(218, nbPlayer, X, nbMove, 173)
        End If
    Next X
End Sub

Sub DoEncore()
    'Count down Encore
    Dim X As Byte
    Dim Y As Byte
    Dim O() As Byte
    O = GetOrderArray
    For Y = 1 To ActNum
        X = O(Y)
        If BCondition(X).Encore Then
            BCondition(X).EncoreDuration = BCondition(X).EncoreDuration - 1
            If BCondition(X).EncoreDuration = 0 Or Current(X).PP(BCondition(X).EncoreMove) = 0 Then
                Call AddMessage(237, nbActive, X)
                BCondition(X).Encore = False
                BCondition(X).EncoreMove = 0
                BCondition(X).EncoreDuration = 0
            End If
        End If
    Next Y
End Sub

Sub DoNightmare()
    Dim X As Byte
    Dim Y As Byte
    Dim O() As Byte
    'Nightmare damage
    O = GetOrderArray
    For Y = 1 To ActNum
        X = O(Y)
        If BCondition(X).Nightmare Then
            If Current(X).Condition = 4 Then
                Call ChangeHP(X, -4)
                Call AddMessage(199, nbActive, X)
                Call SyncPKMN
                If BattleMode = nbAdvBattle And NeedSwitch Then Exit For
            Else
                BCondition(X).Nightmare = False
            End If
        End If
    Next Y
End Sub
Sub DoPerishSong()
    'Count down Perish Song's effect
    Dim X As Byte
    Dim Y As Byte
    Dim O() As Byte
    O = GetOrderArray
    For Y = 1 To ActNum
        X = O(Y)
        If Not BCondition(X).DoneCheck Then
            BCondition(X).DoneCheck = True
            If BCondition(X).PerishSong > 0 Then
               BCondition(X).PerishSong = BCondition(X).PerishSong - 1
                Call AddMessage(248, nbActive, X, nbNumber, BCondition(X).PerishSong)
                If BCondition(X).PerishSong = 0 Then
                    Current(X).HP = 0
                    Call SyncPKMN
                    If BattleMode = nbAdvBattle And NeedSwitch Then
                        If FaintedPokes(TeamNum(X)) <> NumPoke Then Exit Sub
                    End If
                End If
            End If
        End If
    Next Y
    For X = 1 To ActNum
        BCondition(X).DoneCheck = False
    Next X
End Sub

Sub DoLeftovers()
    Dim X As Byte
    Dim Y As Byte
    Dim O() As Byte
    O = GetOrderArray
    For Y = 1 To ActNum
        X = O(Y)
        If Current(X).HP > 0 And Current(X).HP < Current(X).MaxHP And Current(X).Item = nbLeftovers Then
            Current(X).HP = Current(X).HP + Current(X).MaxHP / 16
            Call AddMessage(130, nbActive, X, nbItem, 21)
            Call SyncPKMN
        End If
    Next Y
End Sub


Public Function CurrentWeather() As WeatherTypes
    Dim X As Long
    For X = 1 To ActNum
        If (Current(X).Attribute = nbAirLock Or Current(X).Attribute = nbCloudNine) And Current(X).HP > 0 Then
            CurrentWeather = nbNoWeather
            Exit Function
        End If
    Next X
    CurrentWeather = RealWeather
End Function
Private Sub KRTest()
    Dim A() As String
    Dim X As Long
    Dim Y As Long
    Dim T As Boolean
    ReDim A(1 To 88)
    A(1) = "Absorb"
    A(2) = "Aeroblast"
    A(3) = "Barrage"
    A(4) = "Beat Up"
    A(5) = "Bide"
    A(6) = "Bonemerang"
    A(7) = "Bone Rush"
    A(8) = "Comet Punch"
    A(9) = "Crabhammer"
    A(10) = "Cross Chop"
    A(11) = "Cut"
    A(12) = "Dig"
    A(13) = "Double-Edge"
    A(14) = "Double Kick"
    A(15) = "Doubleslap"
    A(16) = "Dragon Rage"
    A(17) = "Drill Peck"
    A(18) = "Egg Bomb"
    A(19) = "Extremespeed"
    A(20) = "Faint Attack"
    A(21) = "False Swipe"
    A(22) = "Flail"
    A(23) = "Fly"
    A(24) = "Frustration"
    A(25) = "Fury Attack"
    A(26) = "Fury Cutter"
    A(27) = "Fury Swipes"
    A(28) = "Giga Drain"
    A(29) = "Hidden Power"
    A(30) = "Hi Jump Kick"
    A(31) = "Horn Attack"
    A(32) = "Hydro Pump"
    A(33) = "Jump Kick"
    A(34) = "Karate Chop"
    A(35) = "Leech Life"
    A(36) = "Mach Punch"
    A(37) = "Magnitude"
    A(38) = "Mega Drain"
    A(39) = "Megahorn"
    A(40) = "Mega Kick"
    A(41) = "Mega Punch"
    A(42) = "Night Shade"
    A(43) = "Outrage"
    A(44) = "Pay Day"
    A(45) = "Peck"
    A(46) = "Petal Dance"
    A(47) = "Pin Missile"
    A(48) = "Pound"
    A(49) = "Present"
    A(50) = "Psywave"
    A(51) = "Pursuit"
    A(52) = "Quick Attack"
    A(53) = "Rage"
    A(54) = "Rapid Spin"
    A(55) = "Razor Leaf"
    A(56) = "Razor Wind"
    A(57) = "Return"
    A(58) = "Reversal"
    A(59) = "Rock Throw"
    A(60) = "Rollout"
    A(61) = "Scratch"
    A(62) = "Seismic Toss"
    A(63) = "Skull Bash"
    A(64) = "Sky Attack"
    A(65) = "Slam"
    A(66) = "Slash"
    A(67) = "Snore"
    A(68) = "Solar Beam"
    A(69) = "Sonic Boom"
    A(70) = "Spike Cannon"
    A(71) = "Strength"
    A(72) = "Struggle"
    A(73) = "Submission"
    A(74) = "Super Fang"
    A(75) = "Surf"
    A(76) = "Swift"
    A(77) = "Tackle"
    A(78) = "Take Down"
    A(79) = "Thief"
    A(80) = "Thrash"
    A(81) = "Triple Kick"
    A(82) = "Twineedle"
    A(83) = "Vice Grip"
    A(84) = "Vine Whip"
    A(85) = "Vital Throw"
    A(86) = "Waterfall"
    A(87) = "Water Gun"
    A(88) = "Wing Attack"
    For X = 1 To 88
        A(X) = CStr(GetMoveNum(A(X)))
        If A(X) = "0" Then Stop
    Next X
    Debug.Print "Case ";
    For X = 1 To 251
        T = False
        For Y = 1 To 88
            If Val(A(Y)) = X Then
                T = True
                Exit For
            End If
        Next Y
        If Moves(X).KingsRock <> T Then
            If Not T Then
                Debug.Print X & ", ";
            End If
        End If
    Next X
End Sub
Public Sub RestoreLog()
    Dim i As Long
    Dim T(3) As AddMessageTypes
    Dim V(3) As Long
    Dim X As Long
    Dim Y As Long
    Dim F As Long
    Dim B As Byte
    Dim Temp As String
    Dim Nick(1 To 4) As String
    Dim PName(1 To 4) As String
    Dim Team(1 To 4) As Long
    For X = 1 To ActNum
        Nick(X) = Current(X).Nickname
        PName(X) = Current(X).Name
        Team(X) = Current(X).TeamNumber
    Next X
    
    F = FreeFile
    Open LogFile For Binary Access Read As #F
    Temp = LogFile
    LogFile = vbNullString
    BattleWindow.Messages.Text = vbNullString
    
    Get #F, , B
    DontTouchMe = B - 48
    TextLog = String$(LOF(F) - 1, vbNullChar)
    Get #F, , TextLog
    Close #F
    SetSourceStringASM StrPtr(TextLog)
    Do While GetBitsLeftASM >= 9
        For X = 0 To 3
            T(X) = nbBlank
            V(X) = 0
        Next X
        
        i = StreamOutASM(9)
        X = StreamOutASM(2)
        For X = 0 To X
            T(X) = StreamOutASM(4)
            Select Case T(X)
            Case nbActive
                Y = StreamOutASM(4)
                If Y > 6 Then
                    V(X) = 2
                    Y = Y - 6
                Else
                    V(X) = 1
                End If
                Current(V(X)).Name = PKMN(V(X), Y).Name
                Current(V(X)).TeamNumber = PKMN(V(X), Y).TeamNumber
                Current(V(X)).Nickname = PKMN(V(X), Y).Nickname
            Case nbMove, nbPoke: V(X) = StreamOutASM(9)
            Case nbPlayer: V(X) = StreamOutASM(1) + 1
            Case nbCond: V(X) = StreamOutASM(3) + 1
            Case nbItem, nbTrait: V(X) = StreamOutASM(7)
            Case nbStat, nbRule: V(X) = StreamOutASM(4)
            Case nbNumber: V(X) = StreamOutASM(10)
            Case nbType: V(X) = StreamOutASM(5)
            End Select
        Next X
        AddMessage i, T(0), V(0), T(1), V(1), T(2), V(2), T(3), V(3)
    Loop
    If DontTouchMe <> 8 And DontTouchMe <> 0 Then
        TextLog = Right$(TextLog, 1)
    Else
        TextLog = vbNullString
    End If
    LogFile = Temp
    For X = 1 To ActNum
        Current(X).Nickname = Nick(X)
        Current(X).Name = PName(X)
        Current(X).TeamNumber = Team(X)
    Next X

End Sub
Public Sub InitLogFile()
    Dim X As Long
    Dim Temp As String
    LogFile = SlashPath & "bl" & MD5(ServerAddress & You.Name) & ".log"
    If FileExists(LogFile) Then
        SetAttr LogFile, vbNormal
        Kill LogFile
    End If
    X = FreeFile
    Open LogFile For Binary Access Write As #X
    Put #X, , "8"
    Close #X
    SetAttr LogFile, vbHidden
End Sub
Public Function ModsInvolved() As Boolean
    ModsInvolved = (PKMN(1, 1).GameVersion = nbModAdv Or PKMN(2, 1).GameVersion = nbModAdv)
End Function
