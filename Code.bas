Attribute VB_Name = "Code"
'--------------------------------------------------------------------------------------
'Pokémon NetBattle
'By "TV's Ian" Murray
'Project begun in mid-2001
'First public release in early 2002 (v0.8.0)
'Project homepage: http://www.tvsian.com/netbattle
'--------------------------------------------------------------------------------------
'File type registration from http://www.vbsquare.com/files/tip458.html
'ModuleMultimedia from
'Everything else is either me or MSDN.
'--------------------------------------------------------------------------------------
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
    Type1 As Elements
    Type2 As Elements
    Attribute As Traits
    PAtt(0 To 1) As Traits
    AttNum As Byte
    NatureNum As Byte
    Color1 As Integer
    Color2 As Integer
    Move(1 To 4) As Integer
    MaxPP(1 To 4) As Byte
    PP(1 To 4) As Byte
    Item As Items
    Condition As Conditions
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
    AdvMoves() As Integer
    ADVTM() As Integer
    AdvBreeding() As Integer
    AdvSpecial() As Integer
    AdvTutor() As Integer
    LFOnly() As Integer
    MoveTutor() As Integer
    MoveLevel() As Byte
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
    RecycleItem As Items
    MarkerNum As Byte
    Illegals(3) As String
    BreedIllegals(3) As String
    EggGroup1 As Byte
    EggGroup2 As Byte
    
    'Rest in Slp/Frz Check
    Resting As Boolean

    'This is used for various berrys that involve randomness
    ItemEffect As Byte

    'This is the only really strange one - it sets to their position in your lineup.
    'Used for copying info from current to team.
    TeamNumber As Byte
    
    'These are all used for Database Modding
    ModAttr(1) As Traits
    TotalAdvMoves As Long
    IllegalMod As String
End Type

'Pokedex Text
Public Type PokeDexInfo
    RedBlue As String
    Yellow As String
    Gold As String
    Silver As String
    Crystal As String
    Ruby As String
    Sapphire As String
End Type

'Future Sight attacks
Public Type FutureSightData
    count As Integer
    AttackPower As Integer
    Level As Integer
    AccuracyMod As Integer
    RandomNumber As Byte
    CHit As Boolean
    FSMove As Byte
End Type

'Stuff that affects the current Pokemon until it is switched
Public Type BattleStuff
    'These will dynamically change values during battle calculation
    AttackChange As Integer
    DefenseChange As Integer
    SpeedChange As Integer
    SAttackChange As Integer
    SDefenseChange As Integer
    EvadeChange As Integer
    AccuracyChange As Integer
    'For use with Thrash/Outrage
    StuckMove As Integer
    StuckCount As Byte
    'Record last 10 moves used
    MoveUsed(1 To 10) As Integer
    'Seeded
    LeechSeed As Byte '<-- needs a target number now
    'Attracted
    Attract As Byte
    'Confused
    Confuse As Boolean
    ConfuseCounter As Byte
    'For Solarbeam, Hyper Beam, etc.
    Charging As Boolean
    Recharging As Boolean
    'Pokemon is out of the playfield
    'May still be hit by certain attacks
'    Dig As Boolean
'    Fly As Boolean
    'Focus Energy
    FocusEnergy As Boolean
    'Pokemon can't be switched
    TrapNum As Byte
    'Pokemon has successfully used Lock-On or Mind Reader
    LockOn As Byte
    LockOnCount As Byte
    'Pokemon has used Foresight - Normal or Fighting can hit Ghosts
    Foresight As Boolean
    'These two work with Bide
    BideDamage As Integer
    BideCount As Byte
    'For hit 2-5 turn moves - Fire Spin, Whirlpool, etc.
    RepeatMove As Integer
    RepeatPos As Byte
    RepeatCount As Byte
    'Pokemon has used Destiny Bond
    DestinyBond As Boolean
    'Counts up Toxic's damage
    ToxicCount As Byte
    'Perish Song Counter
    PerishSong As Byte
    'If the opponent is hit with Nightmare
    Nightmare As Boolean
    'For Rollout
    DefenseCurl As Boolean
    Rollout As Byte
    FuryCutter As Byte
    Curse As Boolean
    'For the Minimize/Stomp combo
    Minimize As Boolean
    'Disable
    DisabledMove As Integer
    DisableCount As Byte
    'Remaining HP for a Substitute
    Substitute As Integer
    'For the three Protect moves
    ProtectPercent As Byte
    'Under the effects of Mist
    Mist As Boolean
    'Encore
    Encore As Boolean
    EncoreMove As Integer
    EncoreDuration As Byte
    MimicedMove As Integer
    RageCounter As Byte
    'Transform
    TransformedTo As Byte
    'RBY Burn and Paralyse Penalties
    BurnPenalty As Boolean
    ParalyzePenalty As Boolean
'    'RBY Wrap, Bind, and Clamp
'    WrapCount As Integer
'    WrapMove As Integer
    'Baton Pass
    BatonPassing As Boolean
    
    'Targetted Pokemon
    TargetNum As Byte
    'Pursuit
    InPursuit As Boolean
    BeingPursued As Boolean
    'Flinch/BackOff
    Flinching As Boolean
    'Protect/Detect
    Protected As Boolean
    'Endure
    Enduring As Boolean
    FutureSight As FutureSightData
    
    'Note that the client does not know most
    'of these; only the server can use them.
    LastPhHitter As Byte    'Last hitter that used a Physical move
    LastSpHitter As Byte    'Last hitter that used a Special move
    LastPhDamage As Integer 'Last Physical damage
    LastSpDamage As Integer 'Last Special damage
    LastAnyHitter As Byte   'Last hitter that did damage, used for Bide
    LastMove As Integer     'Last move this pokemon used
    LastMoveSlot As Byte    'Slot of the last move this pokemon used
    LastAttack As Integer   'Last attack used against this pokemon
    DBPos As Byte           'Poke that will be affected by Destiny Bond
    MirrorPos As Byte       'Last person that hit this pokemon
    MirrorMove As Integer   'Attack to be used for Mirror Move
    SketchSlot As Byte      'Slot containing the move to be Sketched
    
    
    '---RUBY/SAPP CONDITIONS---
    'If for the Double Battle targeting system
    Targeted As Boolean
    'For Revenge and Focus Punch
    Damaged As Boolean
    HelpingHand As Boolean
    Stockpile As Byte
    YawnCount As Byte
    YawnSleepDuration As Byte
    'Charge
    ChargeCount As Byte
    WishCount As Byte
    'The Pokemon that made the wish
    WishPos As Byte
    Grudge As Boolean
    Ingrain As Boolean
    'Fly/Dig/Dive/Bounce
    SemiInvul As SemiInvulTypes
    Truant As Boolean
    ShedSkin As Boolean
    Snatch As Byte
    MagicCoat As Boolean
    FlashFire As Boolean
    'Use for Future Sight and Perish Song
    DoneCheck As Boolean
    SwitchedThisTurn As Boolean
    TraceNumber As Byte
    TraceOK As Boolean
    OutrageConfuse As Byte
    'For Lansat's Crit raising power
    LansatBerry As Boolean
    'Stores if Fake Out is ok or not
    FakeOut As Byte
    WaterSport As Boolean
    MudSport As Boolean
    Imprison As Boolean
    SealedMove(1 To 4) As Long
    Torment As Boolean
    TauntCount As Byte
    ChoiceBandSlot As Byte
    FaintLast As Boolean
    FollowMe As Boolean
    'Soundproof is extremely freaky, and needs a variable
    Soundproofed As Byte
    'RBY Explosion/Substitute bizarreness
    ExplosionHP As Integer
End Type

'Things that affect the whole team
Public Type TeamCond
    SafeguardCount As Byte
    ReflectCount As Byte
    LightScreenCount As Byte
    MistCount As Byte
    Spikes As Byte
End Type

'Move stuff
Public Type Move
    'Need to use this one for IconList funkiness
    ID As Integer
    Name As String
    Type As Elements
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
    Target As MoveTargets
    MagicCoat As Boolean
End Type

'Trainer info
Public Type Trainer
    Picture As Byte
    Name As String
    'Decides which set of graphics to use for your own Pokemon
    Version As GFXModes
    ProgVersion As String
    'Extra comes up as a tooltip on the battle screen, and in the challenge window.
    Extra As String
    WinMess As String
    LoseMess As String
End Type

Public Type ServerPKMNData
    Item As Byte
    Level As Byte
    Gender As Byte
    Nickname As String
    AttNum As Byte
    NatureNum As Byte
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
    UnownLetter As Byte
    Shiny As Boolean
    Move(1 To 4) As Integer
End Type

'Master Server Player
'Formerly Matching Server Player, until I changed the server.
Public Type MSPlayer
    Active As Boolean
    Name As String
    SID As String
    Extra As String
    Address As String
    DNSAddress As String
    Authority As Integer
    Picture As Integer
    Version As String
    GFXVer As Integer
    TeamString As String
    TeamChecksum As String
    PokeData(1 To 6) As ServerPKMNData
    SkipXOR As Boolean
    PKMN(1 To 6) As Integer
    PKMNImage(1 To 6) As String
    BattlingWith As Integer
    Rank As String
    Compatibility(0 To 6) As Boolean
    Wins As Long
    Losses As Long
    Ties As Long
    Disconnect As Long
    Unrated As Boolean
    PingTime As Single
    Speed As String
    Ignore() As Integer
    BattleID As Integer
    ShowTeam As Boolean
    PongCount As Integer
    FloodCount As Integer
    CommandLock As String
    KickTimer As Byte
    LockDown As Boolean
    ChangingTeams As Boolean
    WatchID As Integer
    StadiumOK As Boolean
    DCReason As String
    GameVersion As CompatModes
    MessageAllow(1 To 10) As Boolean
    ModHash As String
End Type

'This one's for the language plugin system
Public Type LPlug
    Text As String
    FileName As String
    HasMoves As Boolean
    HasPKMN As Boolean
    HasPDEX As Boolean
    HasBattle As Boolean
    HasProgram As Boolean
    HasMisc As Boolean
End Type

Public Type GraphicData
    FileName() As String
    ByteCount() As Long
    ByteStart() As Long
    InFile() As Byte
End Type

Public Type NatureType
    StatChg(1 To 5) As Integer
    Name As String
End Type

'Debug variable
'Double-click on your Pokemon to toggle during battle
'Type a capital D on the loader window to activate there.
Global DebugMode As Boolean

'If True, the program is running within Visual Basic.
'If False, it's running compiled.
Global InVBMode As Boolean

'BasePKMN = default values for Max Gene Pokemon
Global BasePKMN() As Pokemon

'Your team, unaffected by battles or anything else.
Global StoredPKMN(1 To 6) As Pokemon

'Swap space for Expert Mode.
'Has to do a Global, because VB doesn't let you do Public user types in a form.
Global ExpertPKMN As Pokemon

Global BoxPKMN() As Pokemon

'Move information
Global Moves() As Move

'Type effectiveness chart - (AttackType,DefendType)
Global BattleMatrix(1 To 17, 1 To 17) As Single

'Trainer info
Global You As Trainer
Global StationID As String
Global PKMN(1 To 6) As Pokemon
Global BattleCond As BattleData

'To transfer info between the registry, network dialog, and battle screen
'IsServer also determines a few things during battle
Global ServerAddress As String
Global ServerRegName As String
Global IsServer As Boolean

'Strings for display purposes
Global Gender(2) As String
Global Weather(4) As String
Global Condition(8) As String
Global Element(0 To 17) As String
Global Item(74) As String
Global ItemDesc(74) As String
Global RuleText(1 To 15) As String
Global RuleToolTip(1 To 15) As String
'Global FlavorText(1 To 200) As String
Global FTextOffset() As Long
Global FTextLen() As Integer
Global AttributeText(77) As String
Global AttributeDesc(77) As String
Global ColorText(10) As String
Global StatName(1 To 10) As String
Global PDexText() As PokeDexInfo
Global EvoMethod(0 To 15) As String
Global ModeText(0 To 14) As String
Global TerrainText(9) As String
Global ProgramText(76) As String
Global Nature(24) As NatureType

'Server stuff, mostly
'Anything that didn't need to be Global is in the forms' declarations
'Might still be able to clean it up, but it might not be worth it.
Global RelayServer As Boolean
Global Player() As MSPlayer
Global IsLoaded(256) As Boolean
Global Chances() As Integer
Global Disconnecting() As Boolean
Global RuleSelected(1 To 15) As Boolean
Global ChallengeMode As Byte
Global SelectedPlayer As Integer
Global Challenge As Boolean
Global YourNumber As Integer
Global GameType As Byte
Global ChallTerrain As Byte
Global ChallengeNumber As Integer
Global ChallengePending As Boolean
Global ICalled As Boolean
Global Ranking As String
Global Battling As Boolean
Global ListenWrong As Boolean
Global BattleTemp As String
Global StoredFileName As String

'Options
Global RecentFiles(1 To 4) As String
Global SoundOption As Integer
Global MusicOption As Integer
Global AnimOption As Integer
Global AutoScan As Integer
Global AskOnUpdate As Integer
Global BMessStyle As Integer
Global AllowViewing As Integer
Global LogPrompt As Integer
Global LogSave As Byte
Global SavedPassword As String
Global FancyText As Boolean
Global SoundFile(0 To 10) As String
Global SoundEnable(0 To 10) As Boolean
Global RecentServer(1 To 100) As String
Global UseAI As Integer
Global LFile(100) As LPlug
Global CurrLang As String
Global GetSpeed As Integer
Global DoMultiPaste As Boolean
Global TBSort As Integer
Global Autosave As Byte
Global ReplayPrompt As Byte
Global TBMode As CompatModes
Global AddLineBreaks As Boolean
Global UseBG As Boolean
Global OldInterface As Boolean
Global HasColGFX As Boolean
Global UseTS As Boolean
Global TSFormat As String
Global ParseURLs As Boolean
Global DisplayLines As Long
Global Autoload As Boolean
Global MsgToggle(1 To 10) As Boolean
Global MoveDelay As Integer
Global UseNicks As Boolean
Global UsePrefix As Boolean
Global VerIcons As Boolean
Global ColorNames As Boolean
Global CSFilter() As String

'DirectX Stuff
'Global UseDX As Boolean
Global Const UseDX As Boolean = False
Global UseHiResTimer As Boolean
Global DeviceGUID As String
Global RefreshRate As Long

'Server options (mostly)
Global ServerMessage As String
Global ServerPassword As String
Global MaxUsers As Integer
Global FloodTolerance As Integer
Global AllowNewUsers As Integer
Global AllowOldVersions As Integer
Global SendTimer As Integer
Global AutoLogging As Integer
Global NumLines As Long
Global WrongListen As Boolean
Global RunningServer As Boolean
Global ServerDesc As String
Global Admin As String
Global ServerName As String
Global PublicServer As Boolean
Global RealIP As String
Global ConnectedToReg As Boolean
Global MaxIPs As Byte
Global DefaultBanMsg As String
Global PurgeDays As Integer
Global UseXOR As Boolean

'For internal use only.
Global DoneLoading As Boolean
Global DBChecksum As String
Global PasswordBoxTitle As String
Global PasswordBoxCaption As String
Global Windir As String
Global SysDir As String
Global DidScan As Boolean
Global DataPath As String
Global SlashPath As String
Global DBPassword As String
Global Compatibility(0 To 7) As Boolean
Global IMWindowArray() As IMWindow
Global IMWindowPlayer() As Integer
Global IMWindowFlash() As Boolean
Global WatchForm(1 To 5) As Battle
Global WatchLoaded(1 To 5) As Boolean
Global GFile As GraphicData
Global MoveBoxNum As Integer
Global FromBox As Integer
Global ToBox As Integer
Global CopyFlag As Boolean
Global SplashScreenUp As Boolean
Global ServerRunning As Boolean
Global LogFileNum As Integer
Global TeamChangeFromMS As Boolean
Global TeamChangeFromTB As Boolean
Global DebugLogName As String
Global CmdReplay As Boolean
Global CmdTeam As Boolean
Global FTextFile As String
Global OpenedAsReplay As Boolean
Global GFXTempFile() As String
Global TempbanDuration As Long
Global SwapBCondition As BattleStuff
Global SwapBTeamCond As TeamCond
Global SwapClassPKMN As Pokemon


'For the True Random section
Enum RndStateEnum
    rReady
    rEmpty
    rQuerying
End Enum
Global UseTrueRnd As Boolean
Global RndByte() As Byte
Global RndBit(1 To 8) As Byte
Global TmpByte() As Byte
Global RndThresh As Long
Global RndCache As Long
Global BitPos As Integer
Global RndGroup As Long
Global RndState As RndStateEnum
Global BadSID As Boolean
Global DBModStr As String
Global DBModHash As String
Global DBModName As String

'My constants - networking and ranking info
'Global Const NetChunkSize = 256
Global Const BaseURL = "http://www.tvsian.com/nbupdate/"
Global Const LowestRank = -593
Global Const HighestRank = 13495
Global Const ADVLowestRank = -613
Global Const ADVHighestRank = 14669
Global Const RegAddress = "server.netbattle.net" '"localhost"  '
Global Const MainPort = 30000 '  CLIENT <--> SERVER
Global Const RegPort = 30001  '  SERVER <--> REGISTRY
Global Const RegPortC = 30002 '  CLIENT <--> REGISTRY
Global Const BetaRel = ""
Global Const DoDebugLogs As Boolean = False

'Various Enumurations to make things look nicer

'-------------
'When making a new battle rule, make sure to insert it into this Enum.
'-------------
Enum RuleEnum
    nbSleep = 1
    nbFreeze
    nbSelfKO
    nbUsePPUps
    nbStadiumMode
    nbRandbat
    nbLevelBalance
    nbTimeout
    nbUnrated
    nbExactHP
    nbNoWatch
    nbPresentRule
End Enum

Enum Elements
    nbNoType
    nbNormal
    nbFire
    nbWater
    nbElectr
    nbGrass
    nbIce
    nbFight
    nbPoison
    nbGround
    nbFlying
    nbPsychc
    nbBug
    nbRock
    nbGhost
    nbDragon
    nbDark
    nbSteel
End Enum

Enum BattleModes
    nbRBYBattle
    nbGSCBattle
    nbAdvBattle
End Enum

Enum Conditions
    nbNrm = 1
    nbPsn
    nbTox
    nbSlp
    nbBrn
    nbPar
    nbFrz
    nbFnt
End Enum

Enum CompatModes
    nbRBYTrade
    nbGSCTrade
    nbTrueRuSa
    nbFullAdvance
    nbModAdv
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

Enum GFXModes
    nbGFXGrn
    nbGFXRB
    nbGFXYlo
    nbGFXGld
    nbGFXSil
    nbGFXRS
    nbGFXLF
    nbGFXEme
    nbGFXCol
    nbGFXSml
End Enum

Enum AttackBases
    nbPysicalBased
    nbSpecialBased
End Enum

Enum SemiInvulTypes
    nbFly = 1
    nbDig
    nbBounce
    nbDive
End Enum

Enum WeatherTypes
    nbNoWeather
    nbRaining
    nbSunny
    nbSandstorm
    nbHailstorm
End Enum

Enum SoundFiles
    nbSoundOpening
    nbSoundSignon
    nbSoundChat
    nbSoundChallenge
    nbMusicOpening
    nbMusicGSC
    nbMusicRBY
    nbMusicVictory
    nbMusicLost
    nbMusicChallenge
    nbMusicRuSa
End Enum

Enum MoveTargets
    nbGlobal
    nbSelectedTarget
    nbBothEnemies
    nbSelfAffecting
    nbTeamAffecting
    nbEveryoneElse
    nbRandomEnemy
    nbReactionTarget
    nbMoveCalled
End Enum

Enum Terrains
    nbShortGrass
    nbTallGrass
    nbVeryTallGrass
    nbOcean
    nbPond
    nbsAnd
    nbCave
    nbMountain
    nbUnderwater
    nbStadium
End Enum

Enum AddMessageTypes
    nbBlank
    nbActive
    nbMove
    nbPlayer
    nbCond
    nbItem
    nbStat
    nbNumber
    nbPoke
    nbRule
    nbTrait
    nbType
End Enum

Enum Items
    nbNoItem
    nbBerry
    nbBerryJuice
    nbBitterBerry
    nbBurntBerry
    nbGoldBerry
    nbIceBerry
    nbMintBerry
    nbMiracleBerry
    nbMysteryBerry
    nbParalyzecureBerry
    nbPoisoncureBerry
    nbBerserkGene
    nbBlackBelt
    nbBlackGlasses
    nbBrightPowder
    nbCharcoal
    nbDragonFang
    nbFocusBand
    nbHardStone
    nbKingsRock
    nbLeftovers
    nbLightBall
    nbLuckyPunch
    nbMagnet
    nbMetalCoat
    nbMetalPowder
    nbMiracleSeed
    nbMysticWater
    nbNevermeltIce
    nbPinkBow
    nbPoisonBarb
    nbPolkadotBow
    nbQuickClaw
    nbScopeLens
    nbSharpBeak
    nbSilverPowder
    nbSoftSand
    nbSpellTag
    nbStick
    nbThickClub
    nbTwistedSpoon
    nbCheriBerry
    nbChestoBerry
    nbPechaBerry
    nbRawstBerry
    nbAspearBerry
    nbLeppaBerry
    nbOranBerry
    nbPersimBerry
    nbLumBerry
    nbSitrusBerry
    nbFigyBerry
    nbIapapaBerry
    nbMagoBerry
    nbWikiBerry
    nbAguavBerry
    nbLiechiBerry
    nbGanlonBerry
    nbSalacBerry
    nbPetayaBerry
    nbApicotBerry
    nbLansatBerry
    nbStarfBerry
    nbChoiceBand
    nbDeepSeaScale
    nbDeepSeaTooth
    nbLaxIncense
    nbMachoBrace
    nbMentalHerb
    nbSeaIncense
    nbShellBell
    nbSilkScarf
    nbSoulDew
    nbWhiteHerb
End Enum

Enum Traits
    nbNoTrait
    nbStench
    nbDrizzle
    nbSpeedBoost
    nbBattleArmor
    nbSturdy
    nbDamp
    nbLimber
    nbSandVeil
    nbStatic
    nbVoltAbsorb
    nbWaterAbsorb
    nbOblivious
    nbCloudNine
    nbCompoundEyes
    nbInsomnia
    nbColorChange
    nbImmunity
    nbFlashFire
    nbShieldDust
    nbOwnTempo
    nbSuctionCups
    nbIntimidate
    nbShadowTag
    nbRoughSkin
    nbWonderGuard
    nbLevitate
    nbEffectSpore
    nbSynchronize
    nbClearBody
    nbNaturalCure
    nbLightningRod
    nbSereneGrace
    nbSwiftSwim
    nbChlorophyll
    nbIlluminate
    nbTrace
    nbHugePower
    nbPoisonPoint
    nbInnerFocus
    nbMagmaArmor
    nbWaterVeil
    nbMagnetPull
    nbSoundproof
    nbRainDish
    nbSandStream
    nbPressure
    nbThickFat
    nbEarlyBird
    nbFlameBody
    nbRunAway
    nbKeenEye
    nbHyperCutter
    nbPickup
    nbTruant
    nbHustle
    nbCuteCharm
    nbPlus
    nbMinus
    nbForecast
    nbStickyHold
    nbShedSkin
    nbGuts
    nbMarvelScale
    nbLiquidOoze
    nbOvergrow
    nbBlaze
    nbTorrent
    nbSwarm
    nbRockHead
    nbDrought
    nbArenaTrap
    nbVitalSpirit
    nbWhiteSmoke
    nbPurePower
    nbShellArmor
    nbCacophony
    nbAirLock
End Enum

'This defines how many bytes are in a Pokemon String converted
'with PKMN2Str.
Public Const POKELEN = 35

'--------------------------------------------------------------------------------------
'Stuff below here comes from sample code from the controls, or code from other sources.
'--------------------------------------------------------------------------------------

'For the Win32 Internet stuff
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0   ' indicates to use config info from registry
Public Const INTERNET_FLAG_EXISITING_CONNECT = &H20000000
Public Const INTERNET_FLAG_RELOAD = &H80000000

'These are for the file registration
Private Const REG_SZ As Long = 1
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const ERROR_SUCCESS = 0
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private PromptOnErr As Boolean
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

'Win32 Internet functions
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'Get the Windows folder (for auto-updater resources)
Public Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Get the system folder
Public Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Flash the taskbar item
Public Declare Function FlashWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
'For generating the station ID
Private Declare Function GetSID Lib "nbGetSid.dll" () As String
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'For tons of stuff
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dest As Any, ByVal numBytes As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (Addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function apiStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Type HostEnt
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type
Public Const WAIT_TIMEOUT = &H102
Public Const AF_INET = 2
Public Const NI_NAMEREQD = 8
Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Public Const WM_SETTEXT = &HC
'For Help support (Requires HTML Help, which should already be on most PCs)
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Const HH_HELP_CONTEXT = &HF
Public Const WM_SETREDRAW = &HB
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_UPDATENOW = &H100
Public Const RDW_INVALIDATE = &H1

'And now, the library of incredibly fast assembly functions written by MasamuneXGP!
Public Declare Function Dec2BinASM Lib "NBAsm.dll" (ByVal X As Long, ByVal Length As Long) As String
Public Declare Function Bin2DecASM Lib "NBAsm.dll" (ByVal pStr As Long) As Long
Public Declare Sub SetSourceStringASM Lib "NBAsm.dll" (ByVal StrPtrOfString As Long)
Public Declare Function StreamOutASM Lib "NBAsm.dll" (ByVal BitLen As Long) As Long
Public Declare Function ReturnParam Lib "NBAsm.dll" (ByVal Param As Long) As Long
Public Declare Function GetParsePosASM Lib "NBAsm.dll" () As Long
Public Declare Function GetBitsLeftASM Lib "NBAsm.dll" () As Long
Public Declare Function GetDNSPtr Lib "NBAsm.dll" (ByVal ServerhWnd As Long) As Long
Public Declare Sub SetDNSIP Lib "NBAsm.dll" (ByVal IP As Long, ByVal PNum As Long)
Public Declare Sub SetDestStringASM Lib "NBAsm.dll" (ByVal VarPtrOfString As Long, ByRef BitPosVariable As Long)
Public Declare Sub StreamInASM Lib "NBAsm.dll" (ByVal X As Long, ByVal F As Long)


'Private Type PivotIndexType
'    Index As Long
'    Value As Long
'End Type
'Private PivotArray() As PivotIndexType

Sub Main()
    Dim TempVar As String
    Dim sPath As String
    Dim sBuf As String
    Dim CSize As Long
    Dim retval As Long
    Dim TempBtl() As String
    Dim Temp As String
    Dim X As Integer
    Dim Y As Long
    Dim Answer As Long
    On Error Resume Next
    Err.Number = 0
    Debug.Print 1 / 0
    InVBMode = (Err.Number <> 0)
    ReDim RTBInfo(0)
    If Right$(App.Path, 1) <> "\" Then
        SlashPath = App.Path & "\"
    Else
        SlashPath = App.Path
    End If
    App.HelpFile = SlashPath & "PokeBattle.chm::/CSHelpText.txt"
    Call OpenDebugLog
    '>>> Call WriteDebugLog("Application Path: " & SlashPath)
    
    If Asc("A") <> 65 Then
        MsgBox "Your computer is using non-English ASCII tables.  Please set your computer's language to English in the Control Panel before using NetBattle." & vbCrLf & "Alternatively, download AppLocale from Microsoft's web site, and use that to launch NetBattle.", vbCritical, "Error"
        End
    End If
    
    If UCase(Command$) = "SERVER" Then
        'If Not InVBMode Then
        '    MsgBox "Sorry, servers are disabled in this version so we can debug the battle code." & vbCrLf & "They should be available again shortly.", vbInformation, "Not Available"
        '    End
        'End If
'        If Not CheckFileAuth(SlashPath & "PokeDB.cdb", 129309, "5BC77F50") _
'        Or Not CheckFileAuth(SlashPath & "MoveDB.cdb", 12983, "11EAC212") _
'        Or Not CheckFileAuth(SlashPath & "TypeDB.cdb", 178, "BD0D44A0") Then
'            MsgBox "The database files are not the correct version for this version of NetBattle or are corrupted." & vbCrLf & "Please reinstall.", vbCritical, "Database Error"
'            End
'        End If
    End If
    
    'Delete any leftover temp files
    If Not App.PrevInstance Then
        ReDim TempBtl(0)
        Temp = Dir(SlashPath & IIf(InVBMode, "*.vbtmp", "*.tmp"), vbHidden)
        While Temp <> ""
            ReDim Preserve TempBtl(UBound(TempBtl) + 1)
            TempBtl(UBound(TempBtl)) = Temp
            Temp = Dir
        Wend
        For X = 1 To UBound(TempBtl)
            Call SetAttr(SlashPath & TempBtl(X), vbNormal)
            Kill SlashPath & TempBtl(X)
        Next X
    End If
    
    On Error GoTo LoadError
    
    Randomize Timer
    SplashScreenUp = False
    DoneLoading = False
    RunningServer = False
    TeamChangeFromMS = False
    OpenedAsReplay = False
    ReDim IMWindowArray(0)
    ReDim IMWindowPlayer(0)
    ReDim IMWindowFlash(0)
    
    For X = 1 To 5
        Set WatchForm(X) = New Battle
        WatchLoaded(X) = False
    Next X
    '>>>Call WriteDebugLog("Initialized window arrays")
    
    DidScan = False
    'Get the Windows folder name
    sBuf = String(255, 0)
    CSize = 255
    retval = GetWindowsDirectoryA(sBuf, CSize)
    sBuf = Left(sBuf, retval)
    Windir = sBuf
    If Right(Windir, 1) <> "\" Then Windir = Windir & "\"
    '>>> Call WriteDebugLog("Windows folder is " & Windir)
    
    'Get the system folder name
    sBuf = String(255, 0)
    CSize = 255
    retval = GetSystemDirectoryA(sBuf, CSize)
    sBuf = Left(sBuf, retval)
    SysDir = sBuf
    If Right(SysDir, 1) <> "\" Then SysDir = SysDir & "\"
    '>>> Call WriteDebugLog("System folder is " & SysDir)
        
    'Build a string out of the version
    If BetaRel <> "" Then
        You.ProgVersion = App.Major & "." & App.Minor & "." & BetaRel
    Else
        You.ProgVersion = App.Major & "." & App.Minor & "." & App.Revision
    End If
    '>>> Call WriteDebugLog("Reported version is " & You.ProgVersion)
    
    'Load the recent files
    RecentFiles(1) = GetSetting("NetBattle", "Recent Files", "1", "")
    RecentFiles(2) = GetSetting("NetBattle", "Recent Files", "2", "")
    RecentFiles(3) = GetSetting("NetBattle", "Recent Files", "3", "")
    RecentFiles(4) = GetSetting("NetBattle", "Recent Files", "4", "")
    
    'This is only used when a team is hidden
    ReDim BasePKMN(1) As Pokemon
    ReDim Moves(0) As Move
    BasePKMN(0).Name = "???"
    
    If FileExists(SlashPath & "gfxcol.bin") Then HasColGFX = True Else HasColGFX = False
    If HasColGFX Then
        If Not CheckFileAuth(SlashPath & "gfxcol.bin", 1321643, "342BBD07") Then MsgBox "The Colosseum image file was detected, but is corrupt." & vbCrLf & "This is probably due to a bad download.", vbExclamation, "File Error": HasColGFX = False
    End If
    
    If Not FileExists(SysDir & "ide21201.vxd") Then
        On Error Resume Next
        Call FileCopy(SlashPath & "ide21201.vxd", SysDir & "ide21201.vxd")
        On Error GoTo LoadError
    End If
    
    'Verify and check the existence of nbGetSid.dll and ide21201.vxd
    'Make sure the serial pulling dll hasn't been tampered with
    Call WriteDebugLog("Starting SID checksum")
'    If Not CheckFileAuth(SlashPath & "nbGetSid.dll", 45056, "7CFF9483") _
'    Or Not CheckFileAuth(SlashPath & "ide21201.vxd", 4720, "F799D2E4") _
'    Or Not CheckFileAuth(SysDir & "ide21201.vxd", 4720, "F799D2E4") Then
    If Not CheckFileAuth(SlashPath & "virtual.drv", 4720, "F799D2E4") Then
        MsgBox "A required system file is missing or corrupted.  Please reinstall the program.", vbCritical, "Error"
        End
    End If
    Call WriteDebugLog("Checksumming complete, getting SID")
    'Get hard drive serial number (Station ID)
 
    StationID = GetSerialNumber
    Call WriteDebugLog("SID is " & DecompressSID(StationID))
    'DEBUG!
    'StationID = "31337"
    
    'MainContainer contains some shared controls and code (to reduce overall filesize).
    'They include imageLists, the graphics loader, and the file dialog.
    'When MainContainer unloads, the program ends.
    MainContainer.Show
    Exit Sub
LoadError:
    MsgBox "An error has occurred.  NetBattle will now exit." & vbCrLf & "Source: " & Err.Source & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical, "Fatal Error"
    '>>> Call WriteDebugLog("Fatal Error")
    If InVBMode Then
        Stop
        Resume
    End If
    End
End Sub

Public Function GetDBChecksum() As String
    'Grabs a total on all the DB entries.
    'For some reason, it doesn't always seem to sync up, so it's
    'not actually checked right now.
    Dim BigDBThingy As Double
    Dim X As Integer
    Dim Y As Integer
    
    BigDBThingy = 0
    For X = 1 To UBound(BasePKMN)
        BigDBThingy = BigDBThingy + BasePKMN(X).MaxHP
        BigDBThingy = BigDBThingy + BasePKMN(X).Attack
        BigDBThingy = BigDBThingy + BasePKMN(X).Defense
        BigDBThingy = BigDBThingy + BasePKMN(X).Speed
        BigDBThingy = BigDBThingy + BasePKMN(X).SpecialAttack
        BigDBThingy = BigDBThingy + BasePKMN(X).SpecialDefense
        BigDBThingy = BigDBThingy + BasePKMN(X).Type1
        BigDBThingy = BigDBThingy + BasePKMN(X).Type2
    Next X
    For X = 1 To 17
        For Y = 1 To 17
            BigDBThingy = BigDBThingy + BattleMatrix(X, Y)
        Next Y
    Next X
    
    GetDBChecksum = BigDBThingy
    If DebugMode Then MsgBox GetDBChecksum, , "DB Checksum"
End Function

'Superceded by MD5
'Public Function EncryptString(ByVal EncryptMe As String) As String
'    'Really stupid one-way string encryption
'    'Used for the user passwords.
'    'Yes, you can theoretically have more than one string get the same result,
'    'But for the minor use it has, it's not important.
'    Dim TempVar As Long
'    Dim TempVar2 As String
'    Dim X As Integer
'
'    For X = 1 To Len(EncryptMe)
'        TempVar = TempVar + (Asc(Mid(EncryptMe, X, 1)) * X)
'    Next
'    If Len(EncryptMe) < 10 Then
'        TempVar2 = "0" & Len(EncryptMe) & TempVar
'    Else
'        TempVar2 = Len(EncryptMe) & TempVar
'    End If
'    EncryptString = TempVar2
'End Function
'
Public Function IsVersionAt(ByVal PVersion As String, ByVal Major As Integer, ByVal Minor As Integer, ByVal Rev As Integer) As Boolean
    'Version checking code.
    'Returns True if the version passed to it (usually an online player) is >= the requested
    Dim MajorVersion As Integer
    Dim MinorVersion As Integer
    Dim Revision As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    '>>> Call WriteDebugLog("IsVersionAt: " & PVersion & " -- " & Major & "." & Minor & "." & Rev)
    On Error Resume Next
    P1 = InStr(1, PVersion, ".")
    P2 = InStr(P1 + 1, PVersion, ".")
    MajorVersion = Val(Left(PVersion, P1 - 1))
    MinorVersion = Val(Mid(PVersion, P1 + 1, P2 - P1 - 1))
    Revision = Val(Right(PVersion, Len(PVersion) - P2))
    
    If MajorVersion < Major Then
        IsVersionAt = False
    ElseIf MinorVersion < Minor Then
        IsVersionAt = False
    ElseIf MinorVersion = Minor And Revision < Rev Then
        IsVersionAt = False
    Else
        IsVersionAt = True
    End If
End Function

Public Function GetPokeRank(ByVal Pokemon As Integer) As Integer
    'Ranks a single Pokemon.
    'Use the BaseValue flag for a Max Gene whatever.
    Dim TempPKMN As Pokemon
    Dim Total As Integer
    Dim X As Integer
    Dim BattleDamage As Single
    Dim MatrixAdjust As Long
    
    TempPKMN = PKMN(Pokemon)
    If TempPKMN.No = 0 Then GetPokeRank = 0: Exit Function
    
    If CompatVersion(TempPKMN.GameVersion) = nbAdvBattle Then
        Total = (GetAdvHP(BasePKMN(TempPKMN.No).BaseHP, TempPKMN.DV_HP, TempPKMN.EV_HP, TempPKMN.Level) / 1.5) _
        + GetAdvStat(BasePKMN(TempPKMN.No).BaseAttack, TempPKMN.DV_Atk, TempPKMN.EV_Atk, TempPKMN.Level, Nature(TempPKMN.NatureNum).StatChg(1)) _
        + GetAdvStat(BasePKMN(TempPKMN.No).BaseDefense, TempPKMN.DV_Def, TempPKMN.EV_Def, TempPKMN.Level, Nature(TempPKMN.NatureNum).StatChg(2)) _
        + GetAdvStat(BasePKMN(TempPKMN.No).BaseSpeed, TempPKMN.DV_Spd, TempPKMN.EV_Spd, TempPKMN.Level, Nature(TempPKMN.NatureNum).StatChg(3)) _
        + GetAdvStat(BasePKMN(TempPKMN.No).BaseSAttack, TempPKMN.DV_SAtk, TempPKMN.EV_SAtk, TempPKMN.Level, Nature(TempPKMN.NatureNum).StatChg(4)) _
        + GetAdvStat(BasePKMN(TempPKMN.No).BaseSDefense, TempPKMN.DV_SDef, TempPKMN.EV_SDef, TempPKMN.Level, Nature(TempPKMN.NatureNum).StatChg(5))
    Else
        Total = (GetHP(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseHP, TempPKMN.DV_HP) / 1.5) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseAttack, TempPKMN.DV_Atk) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseDefense, TempPKMN.DV_Def) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseSAttack, TempPKMN.DV_SAtk) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseSDefense, TempPKMN.DV_SAtk) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseSpeed, TempPKMN.DV_Spd)
    End If
    MatrixAdjust = 0
    For X = 1 To 17
        If BasePKMN(TempPKMN.No).Type2 = 0 Then
            BattleDamage = BattleMatrix(X, BasePKMN(TempPKMN.No).Type1)
        Else
            BattleDamage = BattleMatrix(X, BasePKMN(TempPKMN.No).Type1) * BattleMatrix(X, BasePKMN(TempPKMN.No).Type2)
        End If
        Select Case BattleDamage
            Case 0
                MatrixAdjust = MatrixAdjust + 75
            Case 0.25
                MatrixAdjust = MatrixAdjust + 50
            Case 0.5
                MatrixAdjust = MatrixAdjust + 25
            Case 2
                MatrixAdjust = MatrixAdjust - 35
            Case 4
                MatrixAdjust = MatrixAdjust - 75
        End Select
    Next
    Total = Total + MatrixAdjust
    'Skew for overly powerful PKMN
    Select Case TempPKMN.No
        'Mewtwo, Lugia, Ho-Oh
        Case 150, 249, 250, 382, 383, 384, 386 To 389
            Total = Total + (TempPKMN.Level * 5)
        'Legendary birds & dogs, Mew, Celebi
        Case 144 To 146, 151, 243 To 246, 251, 379 To 381, 385
            Total = Total + (TempPKMN.Level * 2.5)
        'Snorlax, Dragonite, Tyranitar
        Case 143, 149, 248
            Total = Total + (TempPKMN.Level * 1.25)
    End Select
            
    GetPokeRank = Total
End Function
Public Sub DoDebugRank()
    'For the debug menu.
    'Ranks all 251 Pokemon and dumps their max and min rank to a CSV file.
    Dim Answer As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim MatrixAdjust As Long
    Dim BattleDamage As Single
    Dim LoVal As Long
    Dim HiVal As Long
    Dim FileNum As Integer
    Dim TempPKMN As Pokemon
    
    Answer = MsgBox("Calculate for Advance?", vbYesNo + vbQuestion, "Mode")
        
    FileNum = FreeFile
    If Answer = vbYes Then
        Open SlashPath & "ranksadv.csv" For Output As #FileNum
    Else
        Open SlashPath & "ranksgsc.csv" For Output As #FileNum
    End If
    Write #FileNum, "Name", "Low", "High"
    For X = 1 To UBound(BasePKMN)
        If Answer = vbYes Then
            TempPKMN = BasePKMN(X)
            TempPKMN.DV_HP = 0
            TempPKMN.DV_Atk = 0
            TempPKMN.DV_Def = 0
            TempPKMN.DV_Spd = 0
            TempPKMN.DV_SAtk = 0
            TempPKMN.DV_SDef = 0
            TempPKMN.EV_HP = 0
            TempPKMN.EV_Atk = 0
            TempPKMN.EV_Def = 0
            TempPKMN.EV_Spd = 0
            TempPKMN.EV_SAtk = 0
            TempPKMN.EV_SDef = 0
            TempPKMN.Level = 1
            TempPKMN.GameVersion = nbAdvBattle
            PKMN(1) = TempPKMN
            LoVal = GetPokeRank(1)
            TempPKMN = BasePKMN(X)
            TempPKMN.DV_HP = 31
            TempPKMN.DV_Atk = 31
            TempPKMN.DV_Def = 31
            TempPKMN.DV_Spd = 31
            TempPKMN.DV_SAtk = 31
            TempPKMN.DV_SDef = 31
            TempPKMN.EV_HP = 255
            TempPKMN.EV_Atk = 255
            TempPKMN.EV_Def = 255
            TempPKMN.EV_Spd = 255
            TempPKMN.EV_SAtk = 255
            TempPKMN.EV_SDef = 255
            TempPKMN.Level = 100
            TempPKMN.GameVersion = nbAdvBattle
            PKMN(1) = TempPKMN
            HiVal = GetPokeRank(1)
            Write #FileNum, BasePKMN(X).Name, LoVal, HiVal
        Else
            If X <= 251 Then
                TempPKMN = BasePKMN(X)
                TempPKMN.DV_HP = 0
                TempPKMN.DV_Atk = 0
                TempPKMN.DV_Def = 0
                TempPKMN.DV_Spd = 0
                TempPKMN.DV_SAtk = 0
                TempPKMN.DV_SDef = 0
                TempPKMN.Level = 1
                TempPKMN.GameVersion = nbGSCBattle
                PKMN(1) = TempPKMN
                LoVal = GetPokeRank(1)
                TempPKMN = BasePKMN(X)
                TempPKMN.DV_HP = 15
                TempPKMN.DV_Atk = 15
                TempPKMN.DV_Def = 15
                TempPKMN.DV_Spd = 15
                TempPKMN.DV_SAtk = 15
                TempPKMN.DV_SDef = 15
                TempPKMN.Level = 100
                TempPKMN.GameVersion = nbGSCBattle
                PKMN(1) = TempPKMN
                HiVal = GetPokeRank(1)
                Write #FileNum, BasePKMN(X).Name, LoVal, HiVal
            End If
        End If
    Next
    Close #FileNum
    MsgBox "Individual totals written to ranks.csv", vbInformation, "Done"
End Sub


Public Function FileExists(ByVal FileName As String) As Boolean
    'Determines if a file exists
    'Used by the auto-updater dowwnloader thingy.
    On Error GoTo Failed
    If Dir(FileName, vbHidden) = "" Then FileExists = False Else FileExists = True
    If Right$(FileName, 1) = "\" Then FileExists = False
    Exit Function
Failed:
    FileExists = False
End Function

Public Function CreateFileAss(Extension As String, FileType As String, FileTypeName As String, Action As String, AppPath As String, Optional Switch As String = "", Optional SetIcon As Boolean = False, Optional DefaultIcon As String, Optional PromptOnError As Boolean = False) As Boolean
'// You may use this code all you want on the condition you keep this simple comment
'// Anyone who improves the code please let me know.
'// Date     : 21/1/2000
'// Author   : Damien McGivern
'// E-Mail   : Damien@Dingo-Delights.co.uk
'// Web Site : www.dingo-delights.co.uk
'// Purpose  : To create file associations with default icons

'// Improved 23/1/200 - New parameters 'Switch', 'PromptOnError', better error handling

'// Parameters
'// Required    Extension       (Str) ie ".exe"
'// Required    FileType        (Str) ie "VB.Form"
'// Required    FileTYpeName    (Str) ie. "Visual Basic Form"
'// Required    Action          (Str) ie. "Open" or "Edit"
'// Required    AppPath         (Str) ie. "C:\Myapp"
'// Optional    Switch          (Str) ie. "/u"                  Default = ""
'// Optional    SetIcon         (Bol)                           Default = False
'// Optional    DefaultIcon     (Str) ie. "C:\Myapp,0"
'// Optional    PromptOnError   (Bol)                           Default = False

'// HOW IT WORKS
'// Extension(Str)   Default = FileType(Str)

'// FileType(Str)    Default = FileTypeName(Str)
'// "DefaultIcon"     Default = DefaultIcon(Str)
'// "shell"
'// Action(Str)
'// "command"   Default = AppPath(Str) & switch(Str) & " %1"
    On Error GoTo ErrorHandler:

    PromptOnErr = PromptOnError

    '// Check that AppPath exists.
    If Dir(AppPath, vbNormal) = "" Then
        If PromptOnError Then MsgBox "The application path '" & AppPath & "' cannot be found.", vbCritical + vbOKOnly, "DLL/OCX Register"
        CreateFileAss = False
        Exit Function
    End If

    Dim ERROR_CHARS As String: ERROR_CHARS = "\/:*?<>|" & Chr(34)
    Dim i As Integer

    If Asc(Extension) <> 46 Then Extension = "." & Extension
    '// Check extension has "." at front

    '// Check for invalid chars within extension
    For i = 1 To Len(Extension)
        If InStr(1, ERROR_CHARS, Mid(Extension, i, 1), vbTextCompare) Then
            If PromptOnError Then MsgBox "The file extension '" & Extension & "' contains an illegal char (\/:*?<>|" & Chr(34) & ").", vbCritical + vbOKOnly, "DLL/OCX Register"

            CreateFileAss = False
            Exit Function
        End If
    Next

    If Switch <> "" Then Switch = " " & Trim(Switch)
   Action = FileType & "\shell\" & Action & "\command"

    Call CreateSubKey(HKEY_CLASSES_ROOT, Extension)        '// Create .xxx key
    Call CreateSubKey(HKEY_CLASSES_ROOT, Action)           '// Create action key

    If SetIcon Then
        Call CreateSubKey(HKEY_CLASSES_ROOT, (FileType & "\DefaultIcon"))                '// Create default icon key

        If DefaultIcon = "" Then
            '// This line of code sets the application's own icon as the default file icon
            Call SetKeyDefault(HKEY_CLASSES_ROOT, FileType & "\DefaultIcon", Trim(AppPath & ",0"))
        Else
            Call SetKeyDefault(HKEY_CLASSES_ROOT, FileType & "\DefaultIcon", Trim(DefaultIcon))

        End If
    End If
   Call SetKeyDefault(HKEY_CLASSES_ROOT, Extension, FileType)                                  '// Set .xxx key default

    Call SetKeyDefault(HKEY_CLASSES_ROOT, FileType, FileTypeName)                               '// Set file type default

    Call SetKeyDefault(HKEY_CLASSES_ROOT, Action, AppPath & Switch & " %1")                     '// Set Command line

    CreateFileAss = True
    Exit Function

ErrorHandler:

    If PromptOnError Then MsgBox "An error occured while attempting to create the file extension '" & Extension & "'.", vbCritical + vbOKOnly, "DLL/OCX Register"
    CreateFileAss = False

End Function

Private Function CreateSubKey(RootKey As Long, NewKey As String) As Boolean
    '// This function creates a new sub key
    Dim hKey As Long, regReply As Long
    regReply = RegCreateKeyEx(RootKey, NewKey, _
         0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, 0&)

    If regReply <> ERROR_SUCCESS Then
        If PromptOnErr Then MsgBox "An error occured while attempting to to create a registry key.", vbCritical + vbOKOnly, "DLL/OCX Register"
        CreateSubKey = False
    Else
        CreateSubKey = True
    End If

    Call RegCloseKey(hKey)
End Function

Private Function SetKeyDefault(RootKey As Long, Address As String, Value As String) As Boolean
    '// This function sets the default value of the key which is always a string
    Dim regReply As Long, hKey As Long
    regReply = RegOpenKeyEx(RootKey, Address, 0, KEY_ALL_ACCESS, hKey)

    If regReply <> ERROR_SUCCESS Then
        If PromptOnErr Then MsgBox "An error occured while attempting to access the registery.", vbCritical + vbOKOnly, "DLL/OCX Register"
        SetKeyDefault = False
        Exit Function
    End If

    Value = Value & Chr(0)

    regReply = RegSetValueExString(hKey, "", 0&, REG_SZ, Value, Len(Value))

    If regReply <> ERROR_SUCCESS Then
        If PromptOnErr Then MsgBox "An error occured while attempting to set key default value.", vbCritical + vbOKOnly, "DLL/OCX Register"
        SetKeyDefault = False
    Else
        SetKeyDefault = True
    End If

    Call RegCloseKey(hKey)
End Function

Public Sub GetRecentServers()
    Dim X As Integer
    Dim Temp As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If Not FileExists(SlashPath & "servlist.txt") Then
        Temp = ""
        Open SlashPath & "servlist.txt" For Output As #FileNum
        Write #FileNum, "hubert.dnsalias.net"
        For X = 1 To 99
            Write #FileNum, Temp
        Next
        Close #FileNum
    End If
    
    Open SlashPath & "servlist.txt" For Input As #FileNum
    For X = 1 To 100
        Input #FileNum, RecentServer(X)
    Next
    Close #FileNum
End Sub

Public Sub UpdateServerList(ByVal LastServer As String)
    Dim X As Integer
    Dim CurrentPosition As Integer
    Dim FileNum As Integer
        
    FileNum = FreeFile
    If LastServer = "" Or UCase(LastServer) = UCase(RecentServer(1)) Then Exit Sub
    CurrentPosition = 0
    For X = 2 To 100
        If UCase(LastServer) = UCase(RecentServer(X)) Then CurrentPosition = X
    Next
    If CurrentPosition = 0 Then
        For X = 99 To 1 Step -1
            RecentServer(X + 1) = RecentServer(X)
        Next
    Else
        For X = CurrentPosition - 1 To 1 Step -1
            RecentServer(X + 1) = RecentServer(X)
        Next
    End If
    RecentServer(1) = LastServer
    Open SlashPath & "servlist.txt" For Output As #FileNum
    For X = 1 To 100
        Write #FileNum, RecentServer(X)
    Next
    Close #FileNum
End Sub

Function GetSerialNumber(Optional UseThisOne As String) As String
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Temp As String
    Dim Build As String
    Dim C As New clsSerialGrabber
    If UseThisOne = "" Then
        Build = UCase(C.GetSerial)
    Else
        Build = UseThisOne
    End If
    Temp = ""
    For X = 1 To Len(Build)
        Y = Asc(Mid$(Build, X, 1))
        If Y >= 48 And Y <= 90 Then
            Temp = Temp & Chr$(Y)
        End If
    Next X
    Build = ""
    If Temp = "" Or Temp = String$(Len(Temp), "0") Then
        WriteDebugLog "SIDG Starting Final Read"
        BadSID = True
        Temp1 = String$(255, vbNullChar)
        Temp2 = String$(255, vbNullChar)
        Res = GetVolumeInformation(Left$(Windir, 3), Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
        Temp = Hex(SerialNum)
    Else
        BadSID = False
    End If
    
    Temp = MD5(Temp)
    Build = nSpace(256)
    Y = 0
    Randomize
    For X = 1 To 128 Step 4
        Y = Int((40 - 1 + 1) * Rnd) + 1
        Mid(Build, X) = Dec2Bin(Val("&H" & Mid$(Temp, Y, 1)), 4)
    Next X
    Build = Left$(Build, 100)
    GetSerialNumber = Bin2Chr(Build)
End Function
Public Function DecompressSID(SID As String, Optional Fake As Boolean = False) As String
    Dim Build As String
    Dim Temp As String
    Dim X As Long
    Dim Y As Long
    Build = Left$(Chr2Bin(SID), 100)
    For X = 1 To 5
        Temp = Temp & Mid$(Build, X * 20, 1)
    Next X
    Build = Temp & Build
    For X = 1 To 21
        Y = Bin2Dec(Mid$(Build, X * 5 - 4, 5))
        Y = Y + IIf(Y > 8, 56, 49)
        Mid$(Build, X, 1) = Chr$(Y)
    Next X
    If Fake Then Mid(Build, 1, 1) = "Y"
    DecompressSID = Left$(Build, 21)
End Function

Public Function Dec(ByVal HexNum As String) As Long
    If HexNum = "" Then HexNum = "0"
    Dec = CLng("&H" & HexNum)
End Function

Public Function LegalMove(Pokemon As Pokemon, Optional IgnoreDVs As Boolean = False) As String
    Dim X As Integer
    Dim Y As Integer
    Dim BreedingMoves As Integer
    Dim GSBreedingMoves As Integer
    Dim RBYMoves As Integer
    Dim RBYConflict As Integer
    Dim GSCConflict As Integer
    Dim InvalidMove As Boolean
    Dim SpecialMoves As Integer
    Dim GSSpecialMoves As Integer
    Dim SurfingPika As Integer
    Dim OddEggPoke As Boolean
    Dim Temp As String
    Dim B As Boolean
    Dim Mv(4) As Integer
    Dim AllMoves() As Boolean
    
    'First, catalog all the moves into an array.
    ReDim AllMoves(UBound(Moves), 1 To 13)
    For X = 1 To UBound(Moves)
        For Y = 1 To 13
            AllMoves(X, Y) = False
        Next
    Next
    With Pokemon
        Select Case .GameVersion
            Case nbTrueRBY
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCSpecial)
            Case nbTrueGSC
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTutor)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCSpecial)
            Case nbRBYTrade, nbGSCTrade
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTutor)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCSpecial)
            Case nbTrueRuSa
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvSpecial)
            Case nbFullAdvance, nbModAdv
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvSpecial)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvTutor)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvFL)
            Case Else
                If InVBMode Then Stop
                LegalMove = "Invalid Input"
        End Select
        
        For X = 1 To UBound(AllMoves, 2)
            AllMoves(0, X) = False
        Next X
        For X = 1 To 4
            Mv(X) = CInt(Pokemon.Move(X))
        Next X
        InvalidMove = False
        BreedingMoves = 0
        GSBreedingMoves = 0
        RBYMoves = 0
        LegalMove = ""
        
        'Invalid moves.
        For X = 1 To 4
            If Mv(X) <> 0 Then
                 InvalidMove = True
                 For Y = 1 To UBound(AllMoves, 2)
                     If AllMoves(Mv(X), Y) Then InvalidMove = False
                 Next Y
                 If InvalidMove Then
                     LegalMove = .Name & " can't learn " & Moves(Mv(X)).Name & ".  There may have been a recent change to the database."
                     Exit Function
                 End If
             End If
        Next X
        
        If CompatVersion(.GameVersion) = nbGSCBattle Then
            Temp = .Illegals(0)
            If .GameVersion = nbTrueGSC Then Temp = Temp & .Illegals(1)
        Else
            Temp = .Illegals(2)
            If .GameVersion = nbTrueGSC Then Temp = Temp & .Illegals(3)
        End If
        If Not DBIllegalCheck(Mv, Temp) Then
            LegalMove = .Name & "'s moveset contains a combination of moves that is not legally obtainable."
            Exit Function
        End If
        If .GameVersion = nbModAdv Then
            If Not DBIllegalCheck(Mv, BasePKMN(.No).IllegalMod) Then
                LegalMove = "The currently loaded Database Mod prohibits this moveset on " & .Name
                Exit Function
            End If
        End If
        
        'Invalid breeding combinations.
        If Not BreedCheck(Pokemon.No, Mv, .GameVersion) Then
            LegalMove = .Name & "'s moveset contains a combination of Egg Moves that is not legally obtainable."
            Exit Function
        End If
        
        'RBY Moves <-/-> GSCBreed, GSCSpecial
        RBYMoves = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbRBYLevel, nbRBYTM) Then
                RBYMoves = Mv(X)
                Exit For
            End If
        Next X
        
        RBYConflict = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbGSCEgg, nbGSCSpecial) Then
                If Not Moves(Mv(X)).RBYMove Then
                    RBYConflict = Mv(X)
                    B = False
                    Exit For
                End If
            End If
        Next X
        
        If RBYConflict <> 0 And RBYMoves <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(RBYMoves).Name & " and " & Moves(RBYConflict).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine RBY Moves and non-RBY Breeding or Special Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        'Only 1 Special Move
        SpecialMoves = 0
        Y = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbAdvSpecial, nbGSCSpecial) Then
                If Y = 0 Then
                    Y = Mv(X)
                Else
                    SpecialMoves = Mv(X)
                    Exit For
                End If
            End If
        Next X
        
        If SpecialMoves <> 0 Then
            If CompatVersion(.GameVersion) = nbAdvBattle Then
                Select Case .No
                Case 96, 97 'DROWZEE hatched with BELLY DRUM and WISH
                    If (SpecialMoves = 14 And Y = 353) Or (SpecialMoves = 353 And Y = 14) Then SpecialMoves = 0
                Case 108 'LICKITUNG hatched with HEAL BELL and WISH
                    If (SpecialMoves = 89 And Y = 353) Or (SpecialMoves = 353 And Y = 89) Then SpecialMoves = 0
                Case 113, 242 'CHANSEY hatched with SWEET SCENT and WISH
                    If (SpecialMoves = 220 And Y = 353) Or (SpecialMoves = 353 And Y = 220) Then SpecialMoves = 0
                Case 102, 103 'EXEGGCUTE hatched with SWEET SCENT and WISH
                    If (SpecialMoves = 220 And Y = 353) Or (SpecialMoves = 353 And Y = 220) Then SpecialMoves = 0
                Case 115 'KANGASKHAN hatched with YAWN and WISH
                    If (SpecialMoves = 354 And Y = 353) Or (SpecialMoves = 353 And Y = 354) Then SpecialMoves = 0
                Case 83 'FARFETCH'D hatched with YAWN and WISH
                    If (SpecialMoves = 354 And Y = 353) Or (SpecialMoves = 353 And Y = 354) Then SpecialMoves = 0
                End Select
            End If
        End If
        If SpecialMoves <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(Y).Name & " and " & Moves(SpecialMoves).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine two or more Special Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        '5th Check: Egg Moves <-/-> Special Moves
        BreedingMoves = 0
        
        'EXCEPTION: Gligar can have Earthquake and {Wing Attack and/or Counter}
        B = False
        If Pokemon.No = 207 Then
            For X = 1 To 4
                If Mv(X) = 55 Then B = True: Exit For
            Next X
        End If
                       
        SpecialMoves = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbGSCSpecial, nbAdvSpecial) Then
                SpecialMoves = Mv(X)
                Exit For
            End If
        Next X
        
        For X = 1 To 4
            If Not Moves(SpecialMoves).RBYMove Then
                If LegalMoveCheck(AllMoves, Mv(X), nbGSCEgg, nbAdvEgg, , , , , , , , True) Then
                    If (Mv(X) <> 248 And Mv(X) <> 34) Or Not B Then
                        BreedingMoves = Mv(X)
                        Exit For
                    End If
                End If
            Else
                If LegalMoveCheck(AllMoves, Mv(X), nbGSCEgg, nbAdvEgg) Then
                    If (Mv(X) <> 248 And Mv(X) <> 34) Or Not B Then
                        BreedingMoves = Mv(X)
                        Exit For
                    End If
                End If
            End If
        Next X
        
        If BreedingMoves <> 0 And SpecialMoves <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(BreedingMoves).Name & " and " & Moves(SpecialMoves).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine Egg Moves with Special Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        If CompatVersion(.GameVersion) <> nbAdvBattle Then
        
            '6th Check: GSC Egg Moves <-/-> Special Moves or RBY Moves
            'It sounds redundant but it's not.  Just trust me. -_-
            GSBreedingMoves = 0
            For X = 1 To 4
                If LegalMoveCheck(AllMoves, Mv(X), nbGSCEgg) Then
                    If Not Moves(Mv(X)).RBYMove Then
                        GSBreedingMoves = Mv(X)
                        Exit For
                    End If
                End If
            Next X
            
            GSCConflict = 0
            For X = 1 To 4
                If LegalMoveCheck(AllMoves, Mv(X), nbRBYLevel, nbRBYTM) Then
                    GSCConflict = Mv(X)
                    Exit For
                End If
            Next X
                    
            If GSBreedingMoves <> 0 And GSCConflict <> 0 Then
                Temp = .Name & " cannot learn both " & Moves(GSBreedingMoves).Name & " and " & Moves(GSCConflict).Name & "." & vbNewLine
                Temp = Temp & "(Cannot combine non-RBY Egg Moves with RBY Moves or Special Moves.)"
                LegalMove = Temp
                Exit Function
            End If
                    
            'Odd Eggs and Dizzy Punch
            Select Case Pokemon.No
            Case 173, 35, 36, 174, 39, 40, 236, 106, 107, 175, 176, 237, 238, 124, 240, 126, 25, 26, 125, 172, 239
                For X = 1 To 4
                    If Mv(X) = 45 Then OddEggPoke = True
                Next
            End Select
            Select Case Pokemon.No
                Case 25, 26, 125, 172, 239
                    If Pokemon.GameVersion = nbGSCTrade Then OddEggPoke = False
            End Select
            If OddEggPoke And BreedingMoves <> 0 Then
                Temp = .Name & " cannot learn both Dizzy Punch and " & Moves(BreedingMoves).Name & "." & vbNewLine
                Temp = Temp & "(Cannot combine Dizzy Punch and Egg Moves on Odd Egg Pokémon.)"
                LegalMove = Temp
                Exit Function
            End If
            With Pokemon
                If Not IgnoreDVs Then
                    If OddEggPoke And Not ((.DV_Atk = 2 And .DV_Def = 10 And .DV_SAtk = 10 And .DV_Spd = 10) _
                    Or (.DV_Atk = 0 And .DV_Def = 0 And .DV_SAtk = 0 And .DV_Spd = 0)) Then
                        LegalMove = "DVs must be either 2/10/10/10 or 0/0/0/0 in order for Dizzy Punch to be on " & .Name
                        Exit Function
                    End If
                End If
            End With
        End If
        
        'GSC Pre-evolution moves <-/-> RBY Moves
        RBYConflict = 0
        Select Case Pokemon.No
        Case 35, 36, 39, 40
            For X = 1 To 4
                If Mv(X) = 25 Or Mv(X) = 100 Or Mv(X) = 219 Then
                    RBYConflict = Mv(X)
                    Exit For
                End If
            Next X
        Case 25, 26, 124
            For X = 1 To 4
                If Mv(X) = 25 Or Mv(X) = 219 Then
                    RBYConflict = Mv(X)
                    Exit For
                End If
            Next X
        Case 130
            For X = 1 To 4
                If Mv(X) = 68 Then
                    RBYConflict = Mv(X)
                    Exit For
                End If
            Next X
        End Select
        If RBYMoves <> 0 And RBYConflict <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(RBYMoves).Name & " and " & Moves(RBYConflict).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine GSC Pre-Evolution Moves with RBY Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        'Marill/Azurill
        If .No = 183 Or .No = 184 Then
            Y = 0
            For X = 1 To 4
                Select Case Mv(X)
                Case 343, 187, 183, 319, 58
                    Y = Mv(X)
                End Select
            Next X
            For X = 1 To 4
                If LegalMoveCheck(AllMoves, Mv(X), nbAdvEgg) And Y > 0 Then
                    Select Case Mv(X)
                    Case 343, 187, 183, 319, 58
                    Case Else: Exit For
                    End Select
                End If
            Next X
            If X = 5 Then X = 0 Else X = Mv(X)
            If X <> 0 And Y <> 0 Then
                Temp = .Name & " cannot learn both " & Moves(X).Name & " and " & Moves(Y).Name & "." & vbNewLine
                Temp = Temp & "(Cannot combine Marill's Egg Moves with Azurill's Egg Moves.)"
                LegalMove = Temp
                Exit Function
            End If
        End If
            
    End With
End Function
Private Function LegalMoveCheck(AllMoves() As Boolean, MoveNum As Integer, C1 As MoveTypes, Optional C2 As MoveTypes, Optional C3 As MoveTypes, Optional C4 As MoveTypes, Optional C5 As MoveTypes, Optional C6 As MoveTypes, Optional C7 As MoveTypes, Optional C8 As MoveTypes, Optional C9 As MoveTypes, Optional IgnoreRBY As Boolean = False) As Boolean
    Dim Build As Boolean
    Dim X As Integer
    Build = False
    On Error Resume Next
    If AllMoves(MoveNum, C1) = True Then Build = True
    If AllMoves(MoveNum, C2) = True And C2 > 0 Then Build = True
    If AllMoves(MoveNum, C3) = True And C3 > 0 Then Build = True
    If AllMoves(MoveNum, C4) = True And C4 > 0 Then Build = True
    If AllMoves(MoveNum, C5) = True And C5 > 0 Then Build = True
    If AllMoves(MoveNum, C6) = True And C6 > 0 Then Build = True
    If AllMoves(MoveNum, C7) = True And C7 > 0 Then Build = True
    If AllMoves(MoveNum, C8) = True And C8 > 0 Then Build = True
    If AllMoves(MoveNum, C9) = True And C9 > 0 Then Build = True
    For X = LBound(AllMoves, 2) To UBound(AllMoves, 2)
        If X <> C1 And X <> C2 And X <> C3 And X <> C4 And X <> C5 And X <> C6 And X <> C7 And X <> C8 And X <> C9 Then
            If AllMoves(MoveNum, X) = True And Not (IgnoreRBY And (X = nbRBYLevel Or X = nbRBYTM)) Then
                Build = False
                'Debug.Print "Multiple Learned Move - " & Moves(MoveNum).Name & " - " & X
            End If
        End If
    Next X
'    If MoveNum > 0 Then
'        Debug.Print "Move check - " & Moves(MoveNum).Name & " - " & Build & " - " & C1 & ", " & C2 & ", " & C3 & ", " & C4 & ", " & C5 & ", " & C6 & ", " & C7 & ", " & C8 & ", " & C9
'    End If
    LegalMoveCheck = Build
End Function

Sub ValidMoveArray(ByRef PKMN As Pokemon, ByRef FillArray() As Boolean, ByVal MoveType As MoveTypes)
    Dim X As Integer
    
    With BasePKMN(PKMN.No)
        Select Case MoveType
            'RBY Moves
            Case nbRBYLevel
                For X = 1 To UBound(.RBYMoves)
                    FillArray(.RBYMoves(X), MoveType) = True
                Next
            'RBY TMs
            Case nbRBYTM
                For X = 1 To UBound(.RBYTM)
                    FillArray(.RBYTM(X), MoveType) = True
                Next
            'GSC Moves
            Case nbGSCLevel
                For X = 1 To UBound(.BaseMoves)
                    FillArray(.BaseMoves(X), MoveType) = True
                Next
            'GSC TMs
            Case nbGSCTM
                For X = 1 To UBound(.MachineMoves)
                    FillArray(.MachineMoves(X), MoveType) = True
                Next
            'GSC Egg Moves
            Case nbGSCEgg
                For X = 1 To UBound(.BreedingMoves)
                    FillArray(.BreedingMoves(X), MoveType) = True
                Next
            'GSC Tutor Moves
            Case nbGSCTutor
                For X = 1 To UBound(.MoveTutor)
                    FillArray(.MoveTutor(X), MoveType) = True
                Next
            'GSC Special Moves
            'We'll use this array for True RBY and True GSC's Stadium/Crystal/Odd Egg moves - don't fill in the normal Special moves.
            Case nbGSCSpecial
                If PKMN.GameVersion <> nbTrueGSC And PKMN.GameVersion <> nbTrueRBY Then
                    For X = 1 To UBound(.SpecialMoves)
                        FillArray(.SpecialMoves(X), MoveType) = True
                    Next
                End If
            'Advance moves
            Case nbAdvLevel
                If PKMN.GameVersion = nbModAdv Then X = UBound(.AdvMoves) Else X = .TotalAdvMoves
                For X = 1 To X
                    FillArray(.AdvMoves(X), MoveType) = True
                Next
            'Advance TMs
            Case nbAdvTM
                For X = 1 To UBound(.ADVTM)
                    FillArray(.ADVTM(X), MoveType) = True
                Next
            'Advance Egg Moves
            Case nbAdvEgg
                For X = 1 To UBound(.AdvBreeding)
                    FillArray(.AdvBreeding(X), MoveType) = True
                Next
            'Advance Tutor Moves
            Case nbAdvTutor
                For X = 1 To UBound(.AdvTutor)
                    FillArray(.AdvTutor(X), MoveType) = True
                Next
            'Advance Special Moves
            Case nbAdvSpecial
                For X = 1 To UBound(.AdvSpecial)
                    FillArray(.AdvSpecial(X), MoveType) = True
                Next
            'Fire/Leaf Moves
            Case nbAdvFL
                For X = 1 To UBound(.LFOnly)
                    FillArray(.LFOnly(X), MoveType) = True
                Next
        End Select
        Select Case PKMN.GameVersion
            Case nbRBYTrade
                'RBY Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).RBYMove Then FillArray(X, MoveType) = False
                Next
            Case nbGSCTrade
                'RBY/GSC Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).GSCMove Then FillArray(X, MoveType) = False
                Next
            Case nbTrueRuSa, nbFullAdvance, nbModAdv
                'Nothing special
            Case nbTrueRBY
                'RBY Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).RBYMove Then FillArray(X, MoveType) = False
                Next
                'I know this part could be optimized, but I want to keep it easy for any future changes we may need.
                If MoveType = nbGSCSpecial Then
                    Select Case .No
                        Case 54, 55
                            FillArray(6, MoveType) = True
                    End Select
                End If
            Case nbTrueGSC
                'RBY/GSC Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).GSCMove Then FillArray(X, MoveType) = False
                Next
                If MoveType = nbGSCSpecial Then
                    Select Case .No
                        Case 83
                            FillArray(12, MoveType) = True
                        Case 147 To 149
                            FillArray(61, MoveType) = True
                        Case 207
                            FillArray(55, MoveType) = True
                        Case 25, 26, 35, 36, 39, 40, 106, 107, 125, 126, 135, 172 To 174, 236 To 239
                            FillArray(45, MoveType) = True
                    End Select
                End If
        End Select
    End With
End Sub

Public Function Bool2Bin(ByVal Bool As Boolean) As Integer
    If Bool Then Bool2Bin = 1 Else Bool2Bin = 0
End Function

Public Function Bin2Bool(ByVal Bin As Integer) As Boolean
    If Bin = 1 Then Bin2Bool = True Else Bin2Bool = False
End Function

Public Function FixedHex(ByVal Number As Long, ByVal Length As Long) As String
    Dim Temp As String
    Temp = Hex(Number)
    If Length < Len(Temp) Then
        If InVBMode Then Stop
        Temp = Right(Temp, Length)
    End If
    Temp = String$(Length - Len(Temp), "0") & Temp
    FixedHex = Temp
End Function

Public Sub ReadBinArray(ByVal Number As Long, ByRef Variable)
    Dim Bit() As Long
    Dim X As Integer
    
    ReDim Bit(UBound(Variable))
    Bit(LBound(Variable)) = 1
    For X = LBound(Variable) + 1 To UBound(Variable)
        Bit(X) = Bit(X - 1) * 2
    Next
    For X = (UBound(Variable)) To LBound(Variable) Step -1
        Number = Number - Bit(X)
        If Number < 0 Then
            Number = Number + Bit(X)
            Variable(X) = False
        Else
            Variable(X) = True
        End If
    Next
End Sub

Public Function MakeBinArray(Variable() As Boolean) As Long
    Dim Bit() As Long
    Dim X As Integer
    Dim Y As Long
    
    ReDim Bit(UBound(Variable))
    Y = 0
    Bit(LBound(Variable)) = 1
    For X = LBound(Variable) + 1 To UBound(Variable)
        Bit(X) = Bit(X - 1) * 2
    Next
    For X = LBound(Variable) To UBound(Variable)
        If Variable(X) Then Y = Y + Bit(X)
    Next
    MakeBinArray = Y
End Function

Public Function Pad(ByVal Original As String, ByVal Length As Integer) As String
    Dim X As Byte
    If Len(Original) > Length Then Original = Left(Original, Length)
    Original = Original & String$(Length - Len(Original), " ")
    Pad = Original
End Function

Public Function GetStat(ByVal Level As Integer, ByVal Base As Integer, ByVal DV As Integer) As Integer
    GetStat = Int(Level * (Base + DV + 31.9) / 50) + 5
End Function

Public Function GetHP(ByVal Level As Integer, ByVal Base As Integer, ByVal DV As Integer) As Integer
    GetHP = Int(Level * (Base + DV + 31.9 + 50) / 50) + 10
End Function
Public Function GetAdvStat(ByVal BaseVal As Integer, ByVal IV As Byte, ByVal EV As Byte, ByVal Level As Byte, ByVal PMod As Integer)
    Dim p As Single
    Select Case PMod
    Case 1: p = 1.1
    Case 0: p = 1
    Case -1: p = 0.9
    End Select
    GetAdvStat = Int(Int((BaseVal * 2 + IV + Int(EV / 4)) * Level / 100 + 5) * p)
End Function
Public Function GetAdvHP(ByVal BaseVal As Integer, ByVal IV As Byte, ByVal EV As Byte, ByVal Level As Byte)
    If BaseVal = 1 Then GetAdvHP = 1: Exit Function 'For Shedinja
    GetAdvHP = Int((BaseVal * 2 + IV + Int(EV / 4)) * Level / 100 + 10 + Level)
End Function
Public Function ShinyDV(ByVal DV_Atk As Integer) As Boolean
    Select Case DV_Atk
        Case 15, 14, 11, 10, 7, 6, 3, 2
            ShinyDV = True
        Case Else
            ShinyDV = False
    End Select
End Function

Public Function ChooseImage(Poke As Pokemon, ByVal Ver As GFXModes, Optional ByVal BackView As Boolean = False, Optional Weather As Byte = 0)
    Dim AttackVar As Integer
    Dim DefenseVar As Integer
    Dim SpeedVar As Integer
    Dim SpecialVar As Integer
    Dim Unown As Integer
    Dim Temp As String
    Dim GFXFile As String
    
    'nbGFXRB  - Red/Blue
    'nbGFXGrn - Green
    'nbGFXYlo - Yellow
    'nbGFXGld - Gold
    'nbGFXSil - Silver
    'nbGFXRS  - Ruby/Sapphire
    'nbGFXLF  - Leaf/Fire
    'nbGFXCol - Colosseum
    'nbGFXSml - Small (Advance Mini Pics)
    
    If Poke.No > 151 And Ver <= nbGFXYlo Then Ver = nbGFXSil
    If Poke.No > 251 And Ver <= nbGFXSil Then Ver = nbGFXRS
    If BackView And Ver = nbGFXSil Then Ver = nbGFXGld
    If Ver = nbGFXLF And Poke.No > 151 Then Ver = nbGFXRS
    If Ver = nbGFXEme And (Poke.No = 201 Or Poke.No = 327 Or Poke.No > 385 Or BackView) Then Ver = nbGFXRS
    If Ver = nbGFXCol And (BackView Or Not HasColGFX Or Poke.No > 386) Then Ver = nbGFXRS
    
    'Set a default ? graphic
    If Poke.Shiny Then GFXFile = "000rss.gif" Else GFXFile = "000rs.gif"
    
    'Special handling for Unowns
    If Poke.No = 201 Then
        If Ver = nbGFXSml Then
            Temp = Chr$(Poke.UnownLetter + 97)
            If Temp = "{" Then Temp = "ep"
            If Temp = "|" Then Temp = "qw"
            GFXFile = "201" & Temp & "_1"
        ElseIf (Ver = nbGFXSil Or Ver = nbGFXGld) And Poke.UnownLetter <= 26 Then
            GFXFile = "201" & Format$(Poke.UnownLetter + 1, "00")
            If BackView Then GFXFile = GFXFile & "b"
            If Poke.Shiny Then GFXFile = GFXFile & "s"
        ElseIf Ver = nbGFXCol Then
            GFXFile = "201" & Format$(Poke.UnownLetter + 1, "00") & "c"
            If Poke.Shiny Then GFXFile = GFXFile + "s"
        Else
            GFXFile = "unown" & Format$(Poke.UnownLetter + 1, "00")
            If BackView Then GFXFile = GFXFile & "b"
            If Poke.Shiny Then GFXFile = GFXFile & "s"
        End If
        GFXFile = GFXFile & ".gif"
    'Special handling for Castform
    ElseIf Poke.No = 351 Then
        If Ver = nbGFXSml Then
            GFXFile = "351_1.gif"
        ElseIf Ver = nbGFXCol And (Weather = 0 Or Weather = 3) Then
            GFXFile = "351c"
            If Poke.Shiny Then GFXFile = GFXFile & "s"
            GFXFile = GFXFile & ".gif"
        Else
            GFXFile = "351rs"
            If BackView Then GFXFile = GFXFile & "b"
            If Weather = 0 Or Weather = 3 Then
                If Poke.Shiny Then GFXFile = GFXFile & "s"
            Else
                GFXFile = GFXFile & CStr(Weather)
            End If
            GFXFile = GFXFile & ".gif"
        End If
    'Otherwise
    Else
        Select Case Ver
            Case nbGFXRB
                GFXFile = Format(Poke.No, "000") & "rb"
                If BackView Then GFXFile = GFXFile & "b"
            Case nbGFXGrn
                GFXFile = Format(Poke.No, "000") & "rg"
                If BackView Then GFXFile = GFXFile & "b"
            Case nbGFXYlo
                GFXFile = Format(Poke.No, "000") & "y"
                If BackView Then GFXFile = GFXFile & "b"
            Case nbGFXGld
                GFXFile = Format(Poke.No, "000") & "g"
                If BackView Then GFXFile = GFXFile & "b"
                If Poke.Shiny Then GFXFile = GFXFile & "s"
            Case nbGFXSil
                GFXFile = Format(Poke.No, "000") & "s"
                If BackView Then GFXFile = GFXFile & "b"
                If Poke.Shiny Then GFXFile = GFXFile & "s"
            Case nbGFXRS
                GFXFile = Format(Poke.No, "000") & "rs"
                If BackView Then GFXFile = GFXFile & "b"
                If Poke.Shiny Then GFXFile = GFXFile & "s"
            Case nbGFXLF
                GFXFile = Format(Poke.No, "000") & "fl"
                If BackView Then GFXFile = GFXFile & "b"
                If Poke.Shiny Then GFXFile = GFXFile & "s"
            Case nbGFXCol
                GFXFile = Format(Poke.No, "000") & "c"
                If Poke.Shiny Then GFXFile = GFXFile & "s"
            Case nbGFXEme
                GFXFile = Format(Poke.No, "000") & "e"
                If Poke.Shiny Then GFXFile = GFXFile & "s"
            Case nbGFXSml
                GFXFile = Format(Poke.No, "000") & "_1"
        End Select
        GFXFile = GFXFile & ".gif"
    End If
    ChooseImage = GFXFile
    '>>> Call WriteDebugLog("Selected image: " & GFXFile)
End Function

Public Function HiddenPowerStrength(ByVal DV_Atk As Integer, ByVal DV_Def As Integer, ByVal DV_Spd As Integer, ByVal DV_SAtk As Integer) As Integer
    Dim Temp As Byte
    Dim Temp2 As Byte
    
    Temp = 0
    If DV_Atk >= 8 Then Temp = Temp + 8
    If DV_Def >= 8 Then Temp = Temp + 4
    If DV_Spd >= 8 Then Temp = Temp + 2
    If DV_SAtk >= 8 Then Temp = Temp + 1
    
    Temp2 = 0
    If DV_SAtk = 2 Or DV_SAtk = 3 Or DV_SAtk = 6 Or DV_SAtk = 7 Or DV_SAtk = 10 Or DV_SAtk = 11 Or DV_SAtk = 14 Or DV_SAtk = 15 Then Temp2 = Temp2 + 2
    If DV_SAtk / 2 <> Int(DV_SAtk / 2) Then Temp2 = Temp2 + 1
    
    HiddenPowerStrength = Int(((Temp * 5) + Temp2) / 2) + 31
End Function

Public Function HiddenPowerType(ByVal DV_Atk As Integer, ByVal DV_Def As Integer) As Integer
    Dim Temp As Byte
    
    Temp = 0
    If DV_Atk = 2 Or DV_Atk = 3 Or DV_Atk = 6 Or DV_Atk = 7 Or DV_Atk = 10 Or DV_Atk = 11 Or DV_Atk = 14 Or DV_Atk = 15 Then Temp = Temp + 8
    If DV_Atk / 2 <> Int(DV_Atk / 2) Then Temp = Temp + 4
    If DV_Def = 2 Or DV_Def = 3 Or DV_Def = 6 Or DV_Def = 7 Or DV_Def = 10 Or DV_Def = 11 Or DV_Def = 14 Or DV_Def = 15 Then Temp = Temp + 2
    If DV_Def / 2 <> Int(DV_Def / 2) Then Temp = Temp + 1
    Select Case Temp
        Case 0
            HiddenPowerType = 7
        Case 1
            HiddenPowerType = 10
        Case 2
            HiddenPowerType = 8
        Case 3
            HiddenPowerType = 9
        Case 4
            HiddenPowerType = 13
        Case 5
            HiddenPowerType = 12
        Case 6
            HiddenPowerType = 14
        Case 7
            HiddenPowerType = 17
        Case 8
            HiddenPowerType = 2
        Case 9
            HiddenPowerType = 3
        Case 10
            HiddenPowerType = 5
        Case 11
            HiddenPowerType = 4
        Case 12
            HiddenPowerType = 11
        Case 13
            HiddenPowerType = 6
        Case 14
            HiddenPowerType = 15
        Case 15
            HiddenPowerType = 16
    End Select
End Function
Public Function HiddenPowerTypeAdv(Poke As Pokemon) As Integer
    Dim X As Integer
    If Poke.DV_HP Mod 2 = 1 Then X = X + 1
    If Poke.DV_Atk Mod 2 = 1 Then X = X + 2
    If Poke.DV_Def Mod 2 = 1 Then X = X + 4
    If Poke.DV_Spd Mod 2 = 1 Then X = X + 8
    If Poke.DV_SAtk Mod 2 = 1 Then X = X + 16
    If Poke.DV_SDef Mod 2 = 1 Then X = X + 32
    X = (X * 15) \ 63
    Select Case X
    Case 0: HiddenPowerTypeAdv = 7 'Fighting
    Case 1: HiddenPowerTypeAdv = 10 'Flying
    Case 2: HiddenPowerTypeAdv = 8 'Poison
    Case 3: HiddenPowerTypeAdv = 9 'Ground
    Case 4: HiddenPowerTypeAdv = 13 'Rock
    Case 5: HiddenPowerTypeAdv = 12 'Bug
    Case 6: HiddenPowerTypeAdv = 14 'Ghost
    Case 7: HiddenPowerTypeAdv = 17 'Steel
    Case 8: HiddenPowerTypeAdv = 2 'Fire
    Case 9: HiddenPowerTypeAdv = 3 'Water
    Case 10: HiddenPowerTypeAdv = 5 'Grass
    Case 11: HiddenPowerTypeAdv = 4 'Electric
    Case 12: HiddenPowerTypeAdv = 11 'Psychic
    Case 13: HiddenPowerTypeAdv = 6 'Ice
    Case 14: HiddenPowerTypeAdv = 15 'Dragon
    Case 15: HiddenPowerTypeAdv = 16 'Dark
    End Select
End Function
Public Function HiddenPowerStrengthAdv(Poke As Pokemon) As Integer
    Dim X As Integer
    If (Poke.DV_HP \ 2) Mod 2 = 1 Then X = X + 1
    If (Poke.DV_Atk \ 2) Mod 2 = 1 Then X = X + 2
    If (Poke.DV_Def \ 2) Mod 2 = 1 Then X = X + 4
    If (Poke.DV_Spd \ 2) Mod 2 = 1 Then X = X + 8
    If (Poke.DV_SAtk \ 2) Mod 2 = 1 Then X = X + 16
    If (Poke.DV_SDef \ 2) Mod 2 = 1 Then X = X + 32
    HiddenPowerStrengthAdv = X * 40 \ 63 + 30
End Function
'Lop [Count] characters off the left of a string, return those.
Public Function ChopString(ByRef Source As String, ByVal count As Integer) As String
    Dim Temp As String
    If Source = "" Then Exit Function
    Temp = Left$(Source, count)
    Source = Right$(Source, Len(Source) - count)
    ChopString = Temp
End Function

'Parses a string without actually editing it.  Use this over
'ChopString when speed is important.
Public Static Function ParseString(ByVal count As Long, Optional ByVal NewString = vbNullString) As String
    Dim Build As String
    Dim Pos As Long
    If NewString <> vbNullString Then
        Build = NewString
        Pos = 0
    End If
    ParseString = Mid$(Build, Pos + 1, count)
    Pos = Pos + count
End Function

Public Function Dequote(Text As String) As String
    Dim Build As String
    Build = Text
    If Right(Build, 1) = Chr(34) Then Build = Left(Build, Len(Build) - 1)
    If Left(Build, 1) = Chr(34) Then Build = Right(Build, Len(Build) - 1)
    Dequote = Build
End Function
Public Function nSpace(ByVal Number As Long) As String
    'And what do you know!  All of a sudden the
    'Space() function doesn't work!
    '
    'THIS is why I HATE 3rd party controls.  Anyone
    'who names a Constant with the same name as a VB
    'Function should be shot and hanged.  Upside down.
    nSpace = String$(Number, " ")
End Function

Public Function LegendaryCheck(ByVal PNum As Long, _
    Optional ByVal JustUber As Boolean = False) As Byte
    
    Dim X As Byte
    Dim Y As Byte
    
    For X = 1 To 6
        If BasePKMN(Player(PNum).PKMN(X)).Uber Then Y = Y + 1
        If BasePKMN(Player(PNum).PKMN(X)).Legendary And Not JustUber Then Y = Y + 1
    Next
    LegendaryCheck = Y
End Function
Public Function BattleOK() As String
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    If You.Name = "" Then
        BattleOK = "Your name has not been filled in."
        Exit Function
    End If
    For X = 1 To 6
        If PKMN(X).No = 0 Then
            BattleOK = "Your team has less than 6 Pokémon."
            Exit Function
        Else
            Z = 0
            For Y = 1 To 4
                If PKMN(X).Move(Y) = 0 Then
                    Z = Z + 1
                End If
            Next Y
            If Z = 4 Then
                BattleOK = PKMN(X).Name & " has no moves!"
                Exit Function
            End If
        End If
    Next X
    BattleOK = ""
End Function

Public Function GetClassPKMN(BClass As BattleData, Team As Byte, PokeNum As Byte) As Pokemon
    Dim Blank As Pokemon
    SwapClassPKMN = Blank
    Call BClass.SetSwapPKMN(Team, PokeNum)
    GetClassPKMN = SwapClassPKMN
    'SwapClassPKMN = Blank
End Function
Public Function GetClassBC(BClass As BattleData, Position As Byte) As BattleStuff
    Dim Blank As BattleStuff
    SwapBCondition = Blank
    Call BClass.SetSwapBC(Position)
    GetClassBC = SwapBCondition
    'SwapBCondition = Blank
End Function
Public Function GetClassTC(BClass As BattleData, Team As Byte) As TeamCond
    Dim Blank As TeamCond
    SwapBTeamCond = Blank
    Call BClass.SetSwapTC(Team)
    GetClassTC = SwapBTeamCond
    'SwapBTeamCond = Blank
End Function

''Pokemon data conversion
'Public Function Pkmn2Str(ByRef PokeData As Pokemon) As String
'    Dim Temp As String
'    Dim X As Byte
'
'    With PokeData
'        Temp = FixedHex(.No, 3)
'        Temp = Temp & Pad(.Image, 12)
'        Temp = Temp & Pad(.Nickname, 10)
'        Temp = Temp & FixedHex(.Attribute, 2)
'        Temp = Temp & FixedHex(.Move(1), 3)
'        Temp = Temp & FixedHex(.Move(2), 3)
'        Temp = Temp & FixedHex(.Move(3), 3)
'        Temp = Temp & FixedHex(.Move(4), 3)
'        Temp = Temp & FixedHex(.MaxPP(1), 2)
'        Temp = Temp & FixedHex(.MaxPP(2), 2)
'        Temp = Temp & FixedHex(.MaxPP(3), 2)
'        Temp = Temp & FixedHex(.MaxPP(4), 2)
'        Temp = Temp & FixedHex(.PP(1), 2)
'        Temp = Temp & FixedHex(.PP(2), 2)
'        Temp = Temp & FixedHex(.PP(3), 2)
'        Temp = Temp & FixedHex(.PP(4), 2)
'        Temp = Temp & FixedHex(.Item, 2)
'        Temp = Temp & FixedHex(.Condition, 2)
'        Temp = Temp & FixedHex(.ConditionCount, 2)
'        Temp = Temp & FixedHex(.MaxHP, 3)
'        Temp = Temp & FixedHex(.HP, 3)
'        Temp = Temp & FixedHex(.DV_HP, 1)
'        Temp = Temp & FixedHex(.Attack, 3)
'        Temp = Temp & FixedHex(.DV_Atk, 1)
'        Temp = Temp & FixedHex(.Defense, 3)
'        Temp = Temp & FixedHex(.DV_Def, 1)
'        Temp = Temp & FixedHex(.Speed, 3)
'        Temp = Temp & FixedHex(.DV_Spd, 1)
'        Temp = Temp & FixedHex(.SpecialAttack, 3)
'        Temp = Temp & FixedHex(.SpecialDefense, 3)
'        Temp = Temp & FixedHex(.DV_SAtk, 1)
'        Temp = Temp & FixedHex(.BaseSpecial, 3)
'        Temp = Temp & FixedHex(.Level, 2)
'        Temp = Temp & FixedHex(.Gender, 1)
'        Temp = Temp & FixedHex(.TeamNumber, 1)
'    End With
'    Pkmn2Str = Temp
'End Function
'
'Public Function Str2Pkmn(ByVal Temp As String) As Pokemon
'    Dim PokeData As Pokemon
'    Dim X As Byte
'
'    With PokeData
'        .No = Dec(ChopString(Temp, 3))
'        PokeData = BasePKMN(.No)
'        .Image = Trim(ChopString(Temp, 12))
'        .Nickname = Trim(ChopString(Temp, 10))
'        .Attribute = Dec(ChopString(Temp, 2))
'        .Move(1) = Dec(ChopString(Temp, 3))
'        .Move(2) = Dec(ChopString(Temp, 3))
'        .Move(3) = Dec(ChopString(Temp, 3))
'        .Move(4) = Dec(ChopString(Temp, 3))
'        .MaxPP(1) = Dec(ChopString(Temp, 2))
'        .MaxPP(2) = Dec(ChopString(Temp, 2))
'        .MaxPP(3) = Dec(ChopString(Temp, 2))
'        .MaxPP(4) = Dec(ChopString(Temp, 2))
'        .PP(1) = Dec(ChopString(Temp, 2))
'        .PP(2) = Dec(ChopString(Temp, 2))
'        .PP(3) = Dec(ChopString(Temp, 2))
'        .PP(4) = Dec(ChopString(Temp, 2))
'        .Item = Dec(ChopString(Temp, 2))
'        .Condition = Dec(ChopString(Temp, 2))
'        .ConditionCount = Dec(ChopString(Temp, 2))
'        .MaxHP = Dec(ChopString(Temp, 3))
'        .HP = Dec(ChopString(Temp, 3))
'        .DV_HP = Dec(ChopString(Temp, 1))
'        .Attack = Dec(ChopString(Temp, 3))
'        .DV_Atk = Dec(ChopString(Temp, 1))
'        .Defense = Dec(ChopString(Temp, 3))
'        .DV_Def = Dec(ChopString(Temp, 1))
'        .Speed = Dec(ChopString(Temp, 3))
'        .DV_Spd = Dec(ChopString(Temp, 1))
'        .SpecialAttack = Dec(ChopString(Temp, 3))
'        .SpecialDefense = Dec(ChopString(Temp, 3))
'        .DV_SAtk = Dec(ChopString(Temp, 1))
'        .BaseSpecial = Dec(ChopString(Temp, 3))
'        .Level = Dec(ChopString(Temp, 2))
'        .Gender = Dec(ChopString(Temp, 1))
'        .TeamNumber = Dec(Temp)
'    End With
'    Str2Pkmn = PokeData
'End Function
'
'Public Function Cond2Str(ByRef Condition As BattleStuff) As String
'    Dim Temp As String
'    Dim X As Byte
'
'    With Condition
'        Temp = Hex(.AttackChange + 6)
'        Temp = Temp & Hex(.DefenseChange + 6)
'        Temp = Temp & Hex(.SpeedChange + 6)
'        Temp = Temp & Hex(.SAttackChange + 6)
'        Temp = Temp & Hex(.SDefenseChange + 6)
'        Temp = Temp & Hex(.EvadeChange + 6)
'        Temp = Temp & Hex(.AccuracyChange + 6)
'        Temp = Temp & FixedHex(.StuckMove, 3)
'        Temp = Temp & FixedHex(.StuckCount, 2)
'        For X = 1 To 10
'            Temp = Temp & FixedHex(.MoveUsed(X), 3)
'        Next
'        Temp = Temp & Bool2Bin(.LeechSeed)
'        Temp = Temp & Bool2Bin(.Attract)
'        Temp = Temp & Bool2Bin(.Confuse)
'        Temp = Temp & FixedHex(.ConfuseCounter, 2)
'        Temp = Temp & Bool2Bin(.Charging)
'        Temp = Temp & Bool2Bin(.Recharging)
''        Temp = Temp & Bool2Bin(.Dig)
''        Temp = Temp & Bool2Bin(.Fly)
'        Temp = Temp & Bool2Bin(.Locked)
'        Temp = Temp & Bool2Bin(.LockOn)
'        Temp = Temp & Bool2Bin(.Foresight)
'        Temp = Temp & FixedHex(.BideDamage, 4)
'        Temp = Temp & FixedHex(.BideCount, 2)
'        Temp = Temp & FixedHex(.RepeatMove, 3)
'        Temp = Temp & FixedHex(.RepeatCount, 2)
'        Temp = Temp & Bool2Bin(.DestinyBond)
'        Temp = Temp & FixedHex(.ToxicCount, 2)
'        Temp = Temp & FixedHex(.PerishSong, 2)
'        Temp = Temp & Bool2Bin(.Nightmare)
'        Temp = Temp & FixedHex(.LastDamage, 4)
'        Temp = Temp & FixedHex(.LastSDamage, 4)
'        Temp = Temp & Bool2Bin(.DefenseCurl)
'        Temp = Temp & FixedHex(.Rollout, 2)
'        Temp = Temp & FixedHex(.FuryCutter, 2)
'        Temp = Temp & Bool2Bin(.Curse)
'        Temp = Temp & Bool2Bin(.Minimize)
'        Temp = Temp & FixedHex(.DisabledMove, 3)
'        Temp = Temp & FixedHex(.DisableCount, 2)
'        Temp = Temp & FixedHex(.Substitute, 4)
'        Temp = Temp & FixedHex(.ProtectPercent, 2)
'        Temp = Temp & Bool2Bin(.Mist)
'        Temp = Temp & Bool2Bin(.Encore)
'        Temp = Temp & FixedHex(.EncoreMove, 3)
'        Temp = Temp & FixedHex(.EncoreDuration, 2)
'        Temp = Temp & FixedHex(.MimicedMove, 3)
'        Temp = Temp & FixedHex(.RageCounter, 2)
'        Temp = Temp & FixedHex(.TransformedTo, 2)
'        Temp = Temp & Bool2Bin(.BatonPassing)
'    End With
'    Cond2Str = Temp
'End Function
'
'Public Function Str2Cond(ByVal Temp As String) As BattleStuff
'    Dim Condition As BattleStuff
'    Dim X As Byte
'
'    With Condition
'        .AttackChange = Dec(ChopString(Temp, 1)) - 6
'        .DefenseChange = Dec(ChopString(Temp, 1)) - 6
'        .SpeedChange = Dec(ChopString(Temp, 1)) - 6
'        .SAttackChange = Dec(ChopString(Temp, 1)) - 6
'        .SDefenseChange = Dec(ChopString(Temp, 1)) - 6
'        .EvadeChange = Dec(ChopString(Temp, 1)) - 6
'        .AccuracyChange = Dec(ChopString(Temp, 1)) - 6
'        .StuckMove = Dec(ChopString(Temp, 3))
'        .StuckCount = Dec(ChopString(Temp, 2))
'        For X = 1 To 10
'            .MoveUsed(X) = Dec(ChopString(Temp, 3))
'        Next
'        .LeechSeed = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Attract = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Confuse = Bin2Bool(Val(ChopString(Temp, 1)))
'        .ConfuseCounter = Dec(ChopString(Temp, 2))
'        .Charging = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Recharging = Bin2Bool(Val(ChopString(Temp, 1)))
''        .Dig = Bin2Bool(Val(ChopString(Temp, 1)))
''        .Fly = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Locked = Bin2Bool(Val(ChopString(Temp, 1)))
'        .LockOn = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Foresight = Bin2Bool(Val(ChopString(Temp, 1)))
'        .BideDamage = Dec(ChopString(Temp, 4))
'        .BideCount = Dec(ChopString(Temp, 2))
'        .RepeatMove = Dec(ChopString(Temp, 3))
'        .RepeatCount = Dec(ChopString(Temp, 2))
'        .DestinyBond = Bin2Bool(Val(ChopString(Temp, 1)))
'        .ToxicCount = Dec(ChopString(Temp, 2))
'        .PerishSong = Dec(ChopString(Temp, 2))
'        .Nightmare = Bin2Bool(Val(ChopString(Temp, 1)))
'        .LastDamage = Dec(ChopString(Temp, 4))
'        .LastSDamage = Dec(ChopString(Temp, 4))
'        .DefenseCurl = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Rollout = Dec(ChopString(Temp, 2))
'        .FuryCutter = Dec(ChopString(Temp, 2))
'        .Curse = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Minimize = Bin2Bool(Val(ChopString(Temp, 1)))
'        .DisabledMove = Dec(ChopString(Temp, 3))
'        .DisableCount = Dec(ChopString(Temp, 2))
'        .Substitute = Dec(ChopString(Temp, 4))
'        .ProtectPercent = Dec(ChopString(Temp, 2))
'        .Mist = Bin2Bool(Val(ChopString(Temp, 1)))
'        .Encore = Bin2Bool(Val(ChopString(Temp, 1)))
'        .EncoreMove = Dec(ChopString(Temp, 3))
'        .EncoreDuration = Dec(ChopString(Temp, 2))
'        .MimicedMove = Dec(ChopString(Temp, 3))
'        .RageCounter = Dec(ChopString(Temp, 2))
'        .TransformedTo = Dec(ChopString(Temp, 2))
'        .BatonPassing = Bin2Bool(Temp)
'    End With
'    Str2Cond = Condition
'End Function
'
'Public Function TC2Str(ByRef TC As TeamCond) As String
'    Dim Temp As String
'
'    With TC
'        Temp = FixedHex(.SafeGuardCount, 2)
'        Temp = Temp & FixedHex(.ReflectCount, 2)
'        Temp = Temp & FixedHex(.LightScreenCount, 2)
'        Temp = Temp & Bool2Bin(.Spikes)
'    End With
'    TC2Str = Temp
'End Function
'
'Public Function Str2TC(ByVal Temp As String) As TeamCond
'    Dim TC As TeamCond
'
'    With TC
'        .SafeGuardCount = Dec(ChopString(Temp, 2))
'        .ReflectCount = Dec(ChopString(Temp, 2))
'        .LightScreenCount = Dec(ChopString(Temp, 2))
'        .Spikes = Bin2Bool(Val(Temp))
'    End With
'    Str2TC = TC
'End Function

Public Function StatChange(ByVal Value As Integer) As Single
    StatChange = ((Abs(Value) + 2) / 2) ^ Sgn(Value)
End Function

Public Function CorrectText(ByVal OriginalText As String, Optional ForName As Boolean = False) As String
    Dim X As Integer
    If OriginalText = "PNB2.0" Then OriginalText = ""
    If ForName Then
        For X = 1 To Len(OriginalText)
            Select Case Asc(Mid$(OriginalText, X, 1))
            Case 65 To 90  'Capital Letters
            Case 97 To 122 'Lowercase Letters
            Case 48 To 57  'Numbers
            Case 32, 38 To 41, 45, 59, 91, 93, 95 'The following characters:  & ( ) [ ] ' ; . _ - [Space]
            Case 46
                If X > 3 Then
                    If Mid$(OriginalText, X - 3, 3) = "www" Then Mid(OriginalText, X) = vbNullChar
                End If
            Case Else
                Mid(OriginalText, X) = vbNullChar
            End Select
        Next X
    End If
    OriginalText = Replace(OriginalText, vbNullChar, vbNullString)
    OriginalText = Replace(OriginalText, Chr$(1), vbNullString)
    OriginalText = Replace(OriginalText, """", vbNullString)
    CorrectText = OriginalText
End Function


Public Function BattleEligible(ByRef CheckPKMN() As Pokemon) As Boolean
    'Determine if you're 100% ready for battle
    Dim NumPoke As Integer
    Dim X As Byte
    Dim Y As Byte
    Dim MoveTotal As Integer

    BattleEligible = True
    If You.Name = "" Then
        MsgBox "You must set up a team first!", vbCritical, "Error"
        BattleEligible = False
        Exit Function
    End If
    NumPoke = 0
    For X = 1 To 6
        If CheckPKMN(X).No > 0 Then NumPoke = NumPoke + 1
    Next
    If NumPoke < 6 Then
        MsgBox "You must set up a team first!", vbCritical, "Error"
        BattleEligible = False
        Exit Function
    End If
    For X = 1 To 6
        MoveTotal = 0
        For Y = 1 To 4
            MoveTotal = MoveTotal + CheckPKMN(X).Move(Y)
        Next
        If MoveTotal = 0 Then
            MsgBox "One of your Pokemon has no moves!", vbCritical, "Error"
            BattleEligible = False
            Exit Function
        End If
    Next
End Function

Public Sub UpdateListings(ByVal FileToUse As String)
    'Actually adjust the file listing based on the most recently opened/saved
    If RecentFiles(1) = FileToUse Then
        Exit Sub
    ElseIf RecentFiles(1) = "" Then
        RecentFiles(1) = FileToUse
    ElseIf RecentFiles(2) = "" And RecentFiles(1) <> FileToUse Then
        RecentFiles(2) = RecentFiles(1)
        RecentFiles(1) = FileToUse
    ElseIf RecentFiles(3) = "" And RecentFiles(2) <> FileToUse And RecentFiles(1) <> FileToUse Then
        RecentFiles(3) = RecentFiles(2)
        RecentFiles(2) = RecentFiles(1)
        RecentFiles(1) = FileToUse
    ElseIf RecentFiles(4) = "" And RecentFiles(3) <> FileToUse And RecentFiles(2) <> FileToUse And RecentFiles(1) <> FileToUse Then
        RecentFiles(4) = RecentFiles(3)
        RecentFiles(3) = RecentFiles(2)
        RecentFiles(2) = RecentFiles(1)
        RecentFiles(1) = FileToUse
    ElseIf RecentFiles(4) = FileToUse Then
        RecentFiles(4) = RecentFiles(3)
        RecentFiles(3) = RecentFiles(2)
        RecentFiles(2) = RecentFiles(1)
        RecentFiles(1) = FileToUse
    ElseIf RecentFiles(3) = FileToUse Then
        RecentFiles(3) = RecentFiles(2)
        RecentFiles(2) = RecentFiles(1)
        RecentFiles(1) = FileToUse
    ElseIf RecentFiles(2) = FileToUse Then
        RecentFiles(2) = RecentFiles(1)
        RecentFiles(1) = FileToUse
    ElseIf RecentFiles(1) <> FileToUse And RecentFiles(2) <> FileToUse And RecentFiles(3) <> FileToUse And RecentFiles(4) <> FileToUse Then
        RecentFiles(4) = RecentFiles(3)
        RecentFiles(3) = RecentFiles(2)
        RecentFiles(2) = RecentFiles(1)
        RecentFiles(1) = FileToUse
    End If
    SaveSetting "NetBattle", "Recent Files", "1", RecentFiles(1)
    SaveSetting "NetBattle", "Recent Files", "2", RecentFiles(2)
    SaveSetting "NetBattle", "Recent Files", "3", RecentFiles(3)
    SaveSetting "NetBattle", "Recent Files", "4", RecentFiles(4)
End Sub

Function CompatCheck(CheckPKMN() As Pokemon, Optional ByVal SinglePKMN = 0) As Long
    Dim X As Byte
    Dim Y As Byte
    Dim TempCheck(0 To 6) As Boolean
    
'    ModeText(0) = "RBY (Trades)"
'    ModeText(1) = "GSC (Trades)"
'    ModeText(2) = "Advance"
'    ModeText(3) = "Advance +"
'    ModeText(4) = "Advance (Mods)"
'    ModeText(5) = "True RBY"
'    ModeText(6) = "True GSC"
    
    For X = 0 To 6
        TempCheck(X) = True
    Next
    
    CompatCheck = 0
    For X = 1 To 6
        If (SinglePKMN = 0 Or SinglePKMN = X) And CheckPKMN(X).No > 0 Then
            If Not CheckPKMN(X).ExistRBY Then TempCheck(0) = False: TempCheck(5) = False
            If Not CheckPKMN(X).ExistGSC Then TempCheck(6) = False
            If Not (CheckPKMN(X).ExistGSC Or CheckPKMN(X).ExistRBY) Then TempCheck(1) = False
            If Not CheckPKMN(X).ExistAdv Then TempCheck(2) = False
            If CheckPKMN(X).Item > 41 Then
                TempCheck(1) = False
                TempCheck(6) = False
            ElseIf Not AdvItem(CheckPKMN(X).Item) Then
                TempCheck(2) = False
                TempCheck(3) = False
            End If
            For Y = 1 To 4
                If CheckPKMN(X).Move(Y) > 0 Then
                    If Not Moves(CheckPKMN(X).Move(Y)).RBYMove Then TempCheck(0) = False: TempCheck(5) = False
                    If Not Moves(CheckPKMN(X).Move(Y)).GSCMove Then TempCheck(1) = False: TempCheck(6) = False
                    'If Not TradeMoveCheck(CheckPKMN(X).No, CheckPKMN(X).Move(Y)) Then TempCheck(2) = False: TempCheck(3) = False
                    If Not TrueRBYCheck(CheckPKMN(X).No, CheckPKMN(X).Move(Y)) Then TempCheck(5) = False
                    If Not TrueGSCCheck(CheckPKMN(X).No, CheckPKMN(X).Move(Y)) Then TempCheck(6) = False
                    If Not TrueAdvCheck(CheckPKMN(X).No, CheckPKMN(X).Move(Y)) Then TempCheck(2) = False
                    If Not TrueAdvPlusCheck(CheckPKMN(X).No, CheckPKMN(X).Move(Y)) Then TempCheck(3) = False
                End If
            Next
        End If
    Next
    CompatCheck = MakeBinArray(TempCheck)
End Function

Public Function TeamRank(Optional ByVal Pokemon As Integer = 0) As String
    Dim X As Integer
    Dim Total As Long
    Dim Temp As String
    
    For X = 1 To 6
        Total = Total + GetPokeRank(X)
        'Debug.Print PKMN(X).Nickname & " - " & GetPokeRank(X) & " (" & Total & ")"
    Next
    
    Total = Total - LowestRank
    If CompatVersion(PKMN(1).GameVersion) = nbAdvBattle Then
        Temp = Str(Cap(Int((Total * 100) / (ADVHighestRank - ADVLowestRank)), 100))
    Else
        Temp = Str(Cap(Int((Total * 100) / (HighestRank - LowestRank)), 100))
    End If
    
    TeamRank = Temp
End Function


Public Function NewIMWindow(ByVal TName As String, ByVal PNum As Integer, ByVal ImageNum As Byte) As Boolean
    Dim X As Integer
    Dim XPos As Single
    Dim YPos As Single
    On Error GoTo NoneLoaded
    
    '>>> Call WriteDebugLog("IM Window requested")
    For X = 1 To UBound(IMWindowArray)
        If IMWindowPlayer(X) = 0 Then Exit For
    Next
    If X > UBound(IMWindowArray) Then
        ReDim Preserve IMWindowArray(X) As IMWindow
        ReDim Preserve IMWindowPlayer(X) As Integer
        ReDim Preserve IMWindowFlash(X) As Boolean
    End If
    On Error GoTo ETrap
    Set IMWindowArray(X) = New IMWindow
    'IMWindowArray(X).Show
    IMWindowArray(X).ThisPlayer = PNum
    IMWindowArray(X).Icon = MainContainer.Trainers.ListImages(ImageNum).Picture
    IMWindowArray(X).Caption = TName
    IMWindowPlayer(X) = PNum
    IMWindowFlash(X) = True
    XPos = 300 * MainContainer.IMWindowList.Buttons.count
    YPos = 300 * MainContainer.IMWindowList.Buttons.count
    If XPos > MainContainer.Width Or YPos > MainContainer.Height Then XPos = 0: YPos = 0
    Call MainContainer.AddTBItem(TName, ImageNum)
    IMWindowArray(X).Move XPos, YPos, 3870, 4100
    'IMWindowArray(X).Chatbox.SetFocus
    Call MasterServer.AddToIMQueue("SHOW:" & X)
    '>>> Call WriteDebugLog("IM Window added")
    Exit Function
NoneLoaded:
    ReDim IMWindowArray(1) As IMWindow
    ReDim IMWindowPlayer(1) As Integer
    ReDim IMWindowFlash(1) As Boolean
    X = 1
    Resume
ETrap:
End Function

Public Sub KillIMWindow(ByVal PNum As Integer)
    Dim WindowNumber As Integer
    '>>> Call WriteDebugLog("IM Window kill request")
    WindowNumber = IMWindowID(PNum)
    If WindowNumber = 0 Then Exit Sub
    IMWindowArray(WindowNumber).ThisPlayer = 0
    Unload IMWindowArray(WindowNumber)
    Set IMWindowArray(WindowNumber) = Nothing
    Call MainContainer.DelTBItem(Player(IMWindowPlayer(WindowNumber)).Name)
    IMWindowPlayer(WindowNumber) = 0
    If UBound(IMWindowArray) = WindowNumber Then
        ReDim Preserve IMWindowArray(WindowNumber - 1) As IMWindow
        ReDim Preserve IMWindowPlayer(WindowNumber - 1) As Integer
        ReDim Preserve IMWindowFlash(WindowNumber - 1) As Boolean
    End If
    '>>> Call WriteDebugLog("IM Window killed")
End Sub

Public Function IMWindowID(ByVal PNum As Integer) As Integer
    Dim X As Integer
    
    On Error GoTo NotLoaded
    For X = 1 To UBound(IMWindowPlayer)
        If IMWindowPlayer(X) = PNum Then
            IMWindowID = X
            Exit Function
        End If
    Next
NotLoaded:
    IMWindowID = 0
End Function

Public Function TradeMoveCheck(ByVal PkNum As Integer, ByVal MoveNum As Integer) As Boolean
    Dim IsMove As Boolean
    Dim X As Integer
     
    IsMove = False
    If UBound(BasePKMN(PkNum).AdvMoves) > 0 Then
        For X = 1 To UBound(BasePKMN(PkNum).AdvMoves)
            If Abs(BasePKMN(PkNum).AdvMoves(X)) = MoveNum Then IsMove = True
        Next
    End If
    If UBound(BasePKMN(PkNum).ADVTM) > 0 Then
        For X = 1 To UBound(BasePKMN(PkNum).ADVTM)
            If Abs(BasePKMN(PkNum).ADVTM(X)) = MoveNum Then IsMove = True
        Next
    End If
    If UBound(BasePKMN(PkNum).AdvBreeding) > 0 Then
        For X = 1 To UBound(BasePKMN(PkNum).AdvBreeding)
            If Abs(BasePKMN(PkNum).AdvBreeding(X)) = MoveNum Then IsMove = True
        Next
    End If
    If UBound(BasePKMN(PkNum).AdvSpecial) > 0 Then
        For X = 1 To UBound(BasePKMN(PkNum).AdvSpecial)
            If Abs(BasePKMN(PkNum).AdvSpecial(X)) = MoveNum Then IsMove = True
        Next
    End If
    TradeMoveCheck = IsMove
End Function

Public Sub ReadBoxPKMN()
    Dim TempPKMN As Pokemon
    Dim Number As Integer
    Dim Data As String
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Dim ByteArray() As Byte
    Dim UCSize As Long
    Dim HSize As Long
    Dim Worked As Variant
    Dim FileNum As Integer
    Dim TmpBox As String
    Dim Temp As String
    Dim BoxVer As String
    Dim TMove() As Integer
    Call WriteDebugLog("Box load started")
    On Error GoTo LoadFailed
    FileNum = FreeFile
    ReDim BoxPKMN(0) As Pokemon
    If Not FileExists(SlashPath & "stored.pbx") Then Exit Sub
    '>>> Call WriteDebugLog("Box exists")
    Open SlashPath & "stored.pbx" For Binary Access Read As #FileNum
    BoxVer = "      "
    Get #FileNum, , BoxVer
    If Left$(BoxVer, 5) = "BOX2." Then
        Temp = String$(LOF(FileNum) - 6, vbNullChar)
        Get #FileNum, , Temp
        Close #FileNum
        While Len(Temp) >= POKELEN
            ReDim Preserve BoxPKMN(UBound(BoxPKMN) + 1)
            X = Dec(ChopString(Temp, 1))
            If BoxVer = "BOX2.0" Then
                Temp = Left$(Temp, 10) & "     " & Right$(Temp, Len(Temp) - 10)
            End If
            BoxPKMN(UBound(BoxPKMN)) = Str2PKMN(ChopString(Temp, POKELEN))
            BoxPKMN(UBound(BoxPKMN)).MarkerNum = X
        Wend
        '>>> Call WriteDebugLog("Box load complete")
    Else
        Close #FileNum
        Open SlashPath & "stored.pbx" For Input As #FileNum
        Input #FileNum, UCSize
        Close #FileNum
        HSize = Len(CStr(UCSize)) + Len(vbCrLf) + 1
        ReDim ByteArray(FileLen(SlashPath & "stored.pbx") - HSize) As Byte
        Open SlashPath & "stored.pbx" For Binary Access Read As #FileNum
        Get #FileNum, HSize, ByteArray()
        Close #FileNum
        '>>> Call WriteDebugLog("Box Decompressed")
        'Debug.Print Chr(ByteArray(LBound(ByteArray))) & "(" & ByteArray(LBound(ByteArray)) & ")", Chr(ByteArray(UBound(ByteArray))) & "(" & ByteArray(UBound(ByteArray)) & ")", UBound(ByteArray)
        Worked = MainContainer.Compressor.DecompressData(ByteArray(), UCSize)
        TmpBox = "box" & FixedHex(Int(Rnd * 65536), 4) & ".tmp"
        Open SlashPath & TmpBox For Binary Access Write As #FileNum
        Put #FileNum, , ByteArray()
        Close #FileNum
        '>>> Call WriteDebugLog("Opening " & TmpBox)
        Open SlashPath & TmpBox For Input As #FileNum
        While Not EOF(FileNum)
            Input #FileNum, Data
            Number = Dec(ChopString(Data, 3))
            If Number >= 1 And Number <= UBound(BasePKMN) Then
                TempPKMN = BasePKMN(Number)
                With TempPKMN
                    .Nickname = Trim(ChopString(Data, 10))
                    .Level = Dec(ChopString(Data, 2))
                    .Item = Dec(ChopString(Data, 2))
                    .Gender = Dec(ChopString(Data, 1))
                    .DV_Atk = Dec(ChopString(Data, 1))
                    .DV_Def = Dec(ChopString(Data, 1))
                    .DV_Spd = Dec(ChopString(Data, 1))
                    .DV_SAtk = Dec(ChopString(Data, 1))
                    .InBox = Dec(ChopString(Data, 1))
                    If .InBox = 0 Then .InBox = 1
                    For Y = 1 To 4
                        .Move(Y) = Dec(ChopString(Data, 3))
                    Next Y
                    .GameVersion = 1
                End With
                ReDim Preserve BoxPKMN(UBound(BoxPKMN) + 1) As Pokemon
                BoxPKMN(UBound(BoxPKMN)) = TempPKMN
            End If
        Wend
        Close #FileNum
        Kill SlashPath & TmpBox
        '>>> Call WriteDebugLog("Box load complete")
    End If
    
'    For X = 1 To UBound(BoxPKMN)
'        Z = 1
'        With BoxPKMN(X)
'            TMove = .Move
'            For Y = 1 To 4
'                .Move(Y) = 0
'            Next Y
'            For Y = 1 To 4
'                .Move(Z) = TMove(Y)
'                If LegalMove(BoxPKMN(X)) <> "" Then
'                    .Move(Z) = 0
'                Else
'                    Z = Z + 1
'                End If
'            Next Y
'        End With
'    Next X
    
    Exit Sub
LoadFailed:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error loading box."
    '>>> Call WriteDebugLog("Box load failed")
    If InVBMode Then
        Stop
        Resume
    End If
    Close #FileNum
    If FileExists(SlashPath & TmpBox) Then Kill SlashPath & TmpBox
End Sub

Public Sub WriteBoxPKMN()
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    Dim ThisPKMN As String
    Dim ByteArray() As Byte
    Dim HeaderBytes() As Byte
    Dim Worked As Boolean
    Dim FileNum As Integer
    Dim TmpBox As String
    Dim Build As String
    
    '>>> Call WriteDebugLog("Box write requested")
    On Error GoTo SaveError
    FileNum = FreeFile
    If UBound(BoxPKMN) = 0 Then
        If FileExists(SlashPath & "stored.pbx") Then Kill SlashPath & "stored.pbx"
        Exit Sub
    End If
    If FileExists(SlashPath & "stored.pbx") Then
        'This makes a backup of the existing stored.pbx
        'so it isn't lost if an error occurs.
        Name SlashPath & "stored.pbx" As SlashPath & "stored.old"
    End If
    Build = "BOX2.1"
    For X = 1 To UBound(BoxPKMN)
        Build = Build & Hex(BoxPKMN(X).MarkerNum)
        Build = Build & PKMN2Str(BoxPKMN(X))
    Next X
    Open SlashPath & "stored.pbx" For Binary Access Write As #FileNum
    Put #FileNum, , Build
    Close #FileNum
    '>>> Call WriteDebugLog("Box save complete")
    If FileExists(SlashPath & "stored.old") Then Kill SlashPath & "stored.old"
    Exit Sub
SaveError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error saving box"
    Call RestoreBox
    '>>> Call WriteDebugLog("Box save error")
    If InVBMode Then
        Stop
        Resume
    End If
End Sub
Private Sub RestoreBox()
    On Error Resume Next
    Kill SlashPath & "stored.pbx"
    Name SlashPath & "stored.old" As SlashPath & "stored.pbx"
    Kill SlashPath & "stored.old" 'In case the rename didn't work for some reason
End Sub


Public Function GetStoredNumber(ByVal Box As Byte, ByVal Position As Integer) As Integer
    Dim X As Integer
    Dim FoundInBox As Integer
    
    For X = 1 To UBound(BoxPKMN)
        If BoxPKMN(X).InBox = Box Then FoundInBox = FoundInBox + 1
        If FoundInBox = Position Then Exit For
    Next
    If X > UBound(BoxPKMN) Then GetStoredNumber = 0: Exit Function
    GetStoredNumber = X
End Function

Public Sub LoadGFXKeys(Optional ByVal SourceFile = "Graphics.bin")
    Dim FileNum As Integer
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim NewCount As Long
    Dim SelGFX As String
    Dim GFXBytes() As Byte
    Dim GFXFile As String
    Dim AltFileArray() As String
    Dim OldCount As Long
    Dim Temp As String
    Dim T As Single
    Dim F As Byte
    'Loads the new format Graphics file
    'Loading an alternate file doesn't work yet, because I'm still trying to figure this out.
    If SourceFile = "Graphics.bin" Then
        ReDim GFXTempFile(0)
    Else
        ReDim Preserve GFXTempFile(UBound(GFXTempFile) + 1)
    End If
    F = UBound(GFXTempFile)
    Call WriteDebugLog("Loading graphics file")
    T = Timer
    FileNum = FreeFile
    Open SlashPath & SourceFile For Binary Access Read As #FileNum
    ReDim GFXBytes(3)
    Get #FileNum, , GFXBytes
    X = GFXBytes(0) * 256 + GFXBytes(1)
    Y = GFXBytes(2)
    Y = Y * 256 + GFXBytes(3)
    Temp = String$(X, vbNullChar)
    Get #FileNum, , Temp
    MainContainer.Compressor.DecompressString Temp, Y
    If SourceFile = "Graphics.bin" Then
        GFile.FileName = Split(Temp, "|")
        OldCount = 1
        X = UBound(GFile.FileName)
        Y = -Int(-((X * 12) / 8))
    Else
        'I'm guessing the real values start at (1), since FileName(0) = ""
        AltFileArray = Split(Temp, "|")
        OldCount = UBound(GFile.FileName) + 1
        ReDim Preserve GFile.FileName(OldCount + UBound(AltFileArray))
        For Z = OldCount To OldCount + UBound(AltFileArray)
            GFile.FileName(Z) = AltFileArray(Z - (OldCount))
        Next
        X = UBound(GFile.FileName)
        Y = -Int(-(((X - OldCount) * 12) / 8))
    End If
    Temp = String$(Y, vbNullChar)
    Get #FileNum, , Temp
    Temp = Chr2Bin(Temp)
    If SourceFile = "Graphics.bin" Then
        ReDim GFile.ByteCount(1 To X)
        ReDim GFile.ByteStart(1 To X)
        ReDim GFile.InFile(1 To X)
        Z = 0
        For Y = OldCount To X
            GFile.ByteStart(Y) = Z + 1
            GFile.ByteCount(Y) = Bin2Dec(Mid$(Temp, ((Y - OldCount + 1) - 1) * 12 + 1, 12))
            Z = Z + GFile.ByteCount(Y)
            GFile.InFile(Y) = F
        Next Y
    Else
        ReDim Preserve GFile.ByteCount(1 To X + 1)
        ReDim Preserve GFile.ByteStart(1 To X + 1)
        ReDim Preserve GFile.InFile(1 To X + 1)
        Z = 0
        For Y = OldCount To X
            GFile.ByteStart(Y + 1) = Z + 1
            GFile.ByteCount(Y + 1) = Bin2Dec(Mid$(Temp, ((Y - OldCount + 1) - 1) * 12 + 1, 12))
            Z = Z + GFile.ByteCount(Y + 1)
            GFile.InFile(Y + 1) = F
        Next Y
    End If
    ReDim GFXBytes(0 To LOF(FileNum) - Seek(FileNum))
    Get #FileNum, , GFXBytes
    Close #FileNum
    MainContainer.Compressor.DecompressData GFXBytes, Z
    GFXTempFile(F) = SlashPath & SourceFile & FileHex & IIf(InVBMode, ".vbtmp", ".tmp")
    Open GFXTempFile(F) For Binary Access Write As #FileNum
    Put #FileNum, , GFXBytes
    Close #FileNum
    Call SetAttr(GFXTempFile(F), vbHidden + vbReadOnly)
    Call WriteDebugLog("Graphics loaded in " & Timer - T & " seconds.")
End Sub

'Creates a random team in GetTeam format
Public Function MakeChallengeTeam(GameVersion As Byte) As String
    Dim TempPKMN(1 To 6) As Pokemon
    Dim Dupe As Boolean
    Dim Build As String
    Dim MaxNum As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim CheckMove As Move
    Dim TempMove(1 To 4) As Integer
    Dim AllMoves() As Boolean
    Dim MovePool() As Integer
    Dim TempPool() As Integer
    Dim StringMove(1 To 4) As String
    Dim Temp As String
    Dim TryNext As Boolean
    Dim A As Integer
    Dim i As Integer
    Dim Z As Long
    
    Select Case GameVersion
        Case 0
            MaxNum = 151
        Case 1
            MaxNum = 251
        Case 2
            MaxNum = 385 '356
    End Select
        
    For X = 1 To 6
        'Number:
        Do
            Dupe = False
            Y = Int(Rnd * MaxNum) + 1
            For Z = 1 To 6
                If TempPKMN(Z).No = Y Or Y = 292 Then Dupe = True
            Next Z
        Loop Until Dupe = False
        TempPKMN(X) = BasePKMN(Y)
        
        With TempPKMN(X)
            'Level
            .Level = .LevelBal
            .Nickname = .Name
            Select Case GameVersion
                Case 0, 1
                    .GameVersion = GameVersion
                Case 2
                    .GameVersion = 3
            End Select
            
            'Item:
            Select Case GameVersion
                Case 0
                    .Item = 0
                Case 1
                    .Item = GetSmallRnd(nbTwistedSpoon, 1)
                Case 2
                    Do
                        .Item = GetSmallRnd(UBound(Item), 1)
                    Loop Until AdvItem(.Item)
            End Select

            'DVs:
            Select Case GameVersion
                Case 0, 1
                    .DV_Atk = GetSmallRnd(15)
                    .DV_Def = GetSmallRnd(15)
                    .DV_Spd = GetSmallRnd(15)
                    .DV_SAtk = GetSmallRnd(15)
                    .DV_HP = 0
                    If .DV_Atk Mod 2 = 1 Then .DV_HP = .DV_HP + 8
                    If .DV_Def Mod 2 = 1 Then .DV_HP = .DV_HP + 4
                    If .DV_Spd Mod 2 = 1 Then .DV_HP = .DV_HP + 2
                    If .DV_SAtk Mod 2 = 1 Then .DV_HP = .DV_HP + 1
                Case 2
                    .DV_Atk = GetSmallRnd(31)
                    .DV_Def = GetSmallRnd(31)
                    .DV_Spd = GetSmallRnd(31)
                    .DV_SAtk = GetSmallRnd(31)
                    .DV_SDef = GetSmallRnd(31)
                    .DV_HP = GetSmallRnd(31)
            End Select
            
            'EVs, Natures, Trait, & Shiny for Advance Pokemon:
            If GameVersion = 2 Then
                'Random nature
                .NatureNum = GetSmallRnd(UBound(Nature))
                'Now do the EVs
                '(Moved to a separate sub because it's kinda messy)
                Call DoCPKMNEVs(TempPKMN(X))
                'Trait - pick the first if there's only one available, random if two
                If .PAtt(1) = 0 Then
                    .AttNum = 0
                Else
                    .AttNum = GetSmallRnd(1, 0)
                End If
                '1 in 100 chance.  Just for looks.
                .Shiny = (GetSmallRnd(100, 1) = 50)
            End If
            
            'Gender (should use the same distributions as wild, unless I messed up)
            Select Case GameVersion
                Case 0
                    .Gender = 0
                Case 1
                    Select Case .PercentFemale
                        Case -1
                            .Gender = 0
                        Case Else
                            If .DV_Atk < .PercentFemale Then .Gender = 2 Else .Gender = 1
                    End Select
                Case 2
                    Select Case .PercentFemale
                        Case -1
                            .Gender = 0
                        Case Else
                            If GetSmallRnd(16, 1) < .PercentFemale Then .Gender = 2 Else .Gender = 1
                    End Select
            End Select

            'Moves: (The tricky part...)
            'Okay, first, make a list of all available moves.
            ReDim AllMoves(UBound(Moves))
            ReDim MovePool(UBound(Moves))
            For Y = 1 To 4
                TempMove(Y) = 0
            Next Y
            'RBY pool if using RBY
            Select Case GameVersion
            Case 0
                For Y = 1 To UBound(.RBYMoves)
                    AllMoves(.RBYMoves(Y)) = True
                Next Y
                For Y = 1 To UBound(.RBYTM)
                    AllMoves(.RBYTM(Y)) = True
                Next Y
            Case 1 'GSC Pool for GSC mode
                For Y = 1 To UBound(.BaseMoves)
                    AllMoves(.BaseMoves(Y)) = True
                Next Y
                For Y = 1 To UBound(.MachineMoves)
                    AllMoves(.MachineMoves(Y)) = True
                Next Y
                For Y = 1 To UBound(.BreedingMoves)
                    AllMoves(.BreedingMoves(Y)) = True
                Next Y
                For Y = 1 To UBound(.SpecialMoves)
                    AllMoves(.SpecialMoves(Y)) = True
                Next Y
                For Y = 1 To UBound(.MoveTutor)
                    AllMoves(.MoveTutor(Y)) = True
                Next Y
                For Y = 1 To UBound(.RBYMoves)
                    AllMoves(.RBYMoves(Y)) = True
                Next Y
                For Y = 1 To UBound(.RBYTM)
                    AllMoves(.RBYTM(Y)) = True
                Next Y
            Case 2
            'Advance pool for Advance (duh)
                For Y = 1 To UBound(.AdvMoves)
                    AllMoves(.AdvMoves(Y)) = True
                Next Y
                For Y = 1 To UBound(.ADVTM)
                    AllMoves(.ADVTM(Y)) = True
                Next Y
                For Y = 1 To UBound(.AdvBreeding)
                    AllMoves(.AdvBreeding(Y)) = True
                Next Y
                For Y = 1 To UBound(.AdvSpecial)
                    AllMoves(.AdvSpecial(Y)) = True
                Next Y
                For Y = 1 To UBound(.AdvTutor)
                    AllMoves(.AdvTutor(Y)) = True
                Next Y
                For Y = 1 To UBound(.LFOnly)
                    AllMoves(.LFOnly(Y)) = True
                Next Y
            End Select
'            'Filter out GSC moves if using RBY
'            If GameVersion = 0 Then
'                For Y = 1 To UBound(AllMoves)
'                    If Not Moves(Y).RBYMove Then AllMoves(Y) = False
'                Next
'            End If
            Z = 0
            For Y = 1 To UBound(AllMoves)
                If AllMoves(Y) Then
                    Z = Z + 1
                    MovePool(Z) = Y
                End If
            Next Y
            ReDim Preserve MovePool(Z)

            'Okay, now we get to the choosing.
            If UBound(MovePool) <= 4 Then
                'First of all, if the Pokemon has less than 4 moves available,
                'then our job is done and that's that.
                For Y = 1 To UBound(MovePool)
                    .Move(Y) = MovePool(Y)
                Next Y
            Else
            'Now for the random choosing.
                i = 0
                For Y = 1 To 4
                    If Y = 1 Then
                        'First move will be a damaging move, if the Pokemon has one.
                        Z = 0
                        ReDim TempPool(UBound(MovePool))
                        For A = 1 To UBound(MovePool)
                            CheckMove = Moves(MovePool(A))
                            If GameVersion = 0 Then CheckMove = ConvertMove(CheckMove, GameVersion)
                            With CheckMove
                                If .Power > 0 And .SpecialEffect <> 101 And .SpecialEffect <> 37 Then
                                    Z = Z + 1
                                    TempPool(Z) = MovePool(A)
                                End If
                            End With
                        Next A
                        ReDim Preserve TempPool(Z)
                        If Z = 0 Then 'No Damage moves exist, go for complete random.
                            Z = MovePool(GetSmallRnd(UBound(MovePool), 1))
                        Else
                            Z = TempPool(GetSmallRnd(UBound(TempPool), 1))
                        End If
                    Else
                        'Otherwise complete random
                        Z = MovePool(GetSmallRnd(UBound(MovePool), 1))
                    End If

                    'Check for dupes
                     For A = 1 To Y - 1
                        If TempMove(A) = Z Then Exit For
                     Next A
                     If A <> Y Then
                         'Do over
                         Y = Y - 1
                     Else
                        TempMove(Y) = Z
                        If LegalMove(TempPKMN(X)) <> "" Then
                            TempMove(Y) = 0
                            Y = Y - 1
                            i = i + 1
                            If i > 50 Then
                                'Something went wrong, start over.
                                For Z = 1 To 4
                                    TempMove(Z) = 0
                                Next Z
                                i = 0
                                Y = 0
                            End If
                        End If
                    End If
                Next Y
    
                'Now sort
                For Y = 1 To 4
                    StringMove(Y) = Moves(TempMove(Y)).Name
                Next Y
                Call SortStringArray(StringMove)
                For Y = 1 To 4
                    .Move(Y) = TempMove(Y)
                    'Debug.Print "PKMN #" & X, "Move #" & Y, Moves(TempMove(Y)).Name & " (" & TempMove(Y) & ")"
                Next
                '>>> Call WriteDebugLog("Poke" & X & "- Complete")
            End If
        End With
    Next X
    Build = ""
    For X = 1 To 6
        Build = Build & PKMN2Str(TempPKMN(X))
    Next
    'Debug.Print Len(Build)
    MakeChallengeTeam = Build
End Function

Private Sub DoCPKMNEVs(ByRef BalancedPKMN As Pokemon)
    Dim BSTotal As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Lowest As Byte
    Dim StatMod(0 To 5) As Single
    Dim EVBalance(0 To 5) As Single
    Dim EVTemp(5) As Byte
    Dim InOrder(5) As Byte
    Dim EVLeft As Integer
    
    'Note that I divide BaseHP by 2 since HP is so far out of whack with the rest of the stats.
    'This isn't always true, but it can be adjusted if needed.
    
    'Masamune's note: Actually, there's no need to divide by 2 since we're going by BASE stats
    'instead of actual stats.  The Base HP is on the same level as the other stats.
    
    With BalancedPKMN
        For X = 1 To 5
            Select Case Nature(.NatureNum).StatChg(X)
                Case 1
                    StatMod(X) = 1.1
                Case 0
                    StatMod(X) = 1
                Case -1
                    StatMod(X) = 0.9
            End Select
        Next
        'Total up base stats
        BSTotal = (.BaseHP * StatMod(0)) + (.BaseAttack * StatMod(1)) + (.BaseDefense * StatMod(2)) + (.BaseSpeed * StatMod(3)) + (.BaseSAttack * StatMod(4)) + (.BaseSDefense * StatMod(5))
            
        'Figure out the percentage each individual stat contributes
        'Masamune's note: If you leave the percentages as Singles, the
        'EV values will be exact and there will never be any overage.
        EVBalance(0) = ((.BaseHP * StatMod(0)) * 100) / BSTotal
        EVBalance(1) = ((.BaseAttack * StatMod(1)) * 100) / BSTotal
        EVBalance(2) = ((.BaseDefense * StatMod(2)) * 100) / BSTotal
        EVBalance(3) = ((.BaseSpeed * StatMod(3)) * 100) / BSTotal
        EVBalance(4) = ((.BaseSAttack * StatMod(4)) * 100) / BSTotal
        EVBalance(5) = ((.BaseSDefense * StatMod(5)) * 100) / BSTotal
        
        EVLeft = 510
        For X = 0 To 5
            EVTemp(X) = CByte(Round((510 * EVBalance(X)) / 100))
            EVLeft = EVLeft - EVTemp(X)
        Next
            
        'Sort 'em out
        For X = 0 To 5
            For Y = 0 To 5
                If EVTemp(InOrder(Y)) >= EVTemp(X) Then Exit For
            Next
            For Y = 4 To Y Step -1
                InOrder(Y + 1) = InOrder(Y)
            Next Y
            InOrder(Y + 1) = X
        Next
            
        'Okay, let's make sure we didn't go too far either way...
        Y = Sgn(EVLeft)
        X = IIf(Y = 1, 0, 5)
        While EVLeft <> 0
            Z = InOrder(X)
            If Not ((EVTemp(Z) = 255 And Y = 1) Or (EVTemp(Z) = 0 And Y = -1)) Then
                EVTemp(Z) = EVTemp(Z) + Y
                EVLeft = EVLeft - Y
            End If
            X = X + Y
            If X = 6 Then X = 0
            If X = -1 Then X = 5
        Wend
            
        'Since we're Ok, do the change.
        .EV_HP = EVTemp(0)
        .EV_Atk = EVTemp(1)
        .EV_Def = EVTemp(2)
        .EV_Spd = EVTemp(3)
        .EV_SAtk = EVTemp(4)
        .EV_SDef = EVTemp(5)
    End With
End Sub


Public Function TrueRBYCheck(ByVal PkNum As Integer, ByVal MoveNum As Integer) As Boolean
    Dim IsMove As Boolean
    Dim X As Integer
     
    IsMove = False
    For X = 1 To UBound(BasePKMN(PkNum).RBYMoves)
        If Abs(BasePKMN(PkNum).RBYMoves(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).RBYTM)
        If Abs(BasePKMN(PkNum).RBYTM(X)) = MoveNum Then IsMove = True
    Next
    'Amnesia Psyduck exception
    If (PkNum = 54 Or PkNum = 55) And MoveNum = 6 Then IsMove = True
    TrueRBYCheck = IsMove
End Function

Public Sub PrepareMode(ByVal Mode As Integer, ByVal PokeNum As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    'Speed up processing by ignoring versions
    Dim SkipRBY As Boolean
    Dim SkipGSC As Boolean
    Dim SkipAdv As Boolean
    '1=First level, 2=Second level, 3=Third level
    'For values: 1=RBY, 2=GSC, 3=Adv
    Dim Priority(3) As Byte
    
    If Mode < 0 Or Mode > 6 Then Exit Sub
    
    Select Case Mode
        Case 0
            SkipRBY = False
            SkipGSC = False
            SkipAdv = True
            Priority(1) = 1
            Priority(2) = 2
            Priority(3) = 0
        Case 1
            SkipRBY = False
            SkipGSC = False
            SkipAdv = True
            Priority(1) = 2
            Priority(2) = 1
            Priority(3) = 0
        Case 2, 3
            SkipRBY = True
            SkipGSC = True
            SkipAdv = False
            Priority(1) = 3
            Priority(2) = 2
            Priority(3) = 1
        Case 4
            SkipRBY = False
            SkipGSC = False
            SkipAdv = False
            Priority(1) = 3
            Priority(2) = 2
            Priority(3) = 1
        Case 5
            SkipRBY = False
            SkipGSC = True
            SkipAdv = True
            Priority(1) = 1
            Priority(2) = 0
            Priority(3) = 0
        Case 6
            SkipRBY = True
            SkipGSC = False
            SkipAdv = True
            Priority(1) = 0
            Priority(2) = 1
            Priority(3) = 0
    End Select
    
    'Reset the moves
    With BasePKMN(PokeNum)
        If Not SkipRBY Then
            For X = 1 To UBound(.RBYMoves)
                If .RBYMoves(X) < 0 Then .RBYMoves(X) = .RBYMoves(X) * -1
            Next
            For X = 1 To UBound(.RBYTM)
                If .RBYTM(X) < 0 Then .RBYTM(X) = .RBYTM(X) * -1
            Next
        End If
        If Not SkipGSC Then
            For X = 1 To UBound(.BaseMoves)
                If .BaseMoves(X) < 0 Then .BaseMoves(X) = .BaseMoves(X) * -1
            Next
            For X = 1 To UBound(.MachineMoves)
                If .MachineMoves(X) < 0 Then .MachineMoves(X) = .MachineMoves(X) * -1
            Next
            For X = 1 To UBound(.BreedingMoves)
                If .BreedingMoves(X) < 0 Then .BreedingMoves(X) = .BreedingMoves(X) * -1
            Next
            For X = 1 To UBound(.SpecialMoves)
                If .SpecialMoves(X) < 0 Then .SpecialMoves(X) = .SpecialMoves(X) * -1
            Next
            For X = 1 To UBound(.MoveTutor)
                If .MoveTutor(X) < 0 Then .MoveTutor(X) = .MoveTutor(X) * -1
            Next
        End If
        If Not SkipAdv Then
            For X = 1 To UBound(.AdvMoves)
                If .AdvMoves(X) < 0 Then .AdvMoves(X) = .AdvMoves(X) * -1
            Next
            For X = 1 To UBound(.ADVTM)
                If .ADVTM(X) < 0 Then .ADVTM(X) = .ADVTM(X) * -1
            Next
            For X = 1 To UBound(.AdvBreeding)
                If .AdvBreeding(X) < 0 Then .AdvBreeding(X) = .AdvBreeding(X) * -1
            Next
            For X = 1 To UBound(.AdvSpecial)
                If .AdvSpecial(X) < 0 Then .AdvSpecial(X) = .AdvSpecial(X) * -1
            Next
            For X = 1 To UBound(.AdvTutor)
                If .AdvTutor(X) < 0 Then .AdvTutor(X) = .AdvTutor(X) * -1
            Next
        End If
        For Z = 1 To 3
            Select Case Priority(Z)
                Case 0
                    'Do nothing
                Case 1
                    'Process RBY
                    For X = 1 To UBound(.RBYMoves)
                        If .RBYMoves(X) > 0 Then
                            For Y = 1 To UBound(.RBYTM)
                                If .RBYMoves(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                            Next
                            If Not SkipGSC Then
                                For Y = 1 To UBound(.BaseMoves)
                                    If .RBYMoves(X) = .BaseMoves(Y) Then .BaseMoves(Y) = .BaseMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MachineMoves)
                                    If .RBYMoves(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.BreedingMoves)
                                    If .RBYMoves(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.SpecialMoves)
                                    If .RBYMoves(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MoveTutor)
                                    If .RBYMoves(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                                Next
                            End If
                            If Not SkipAdv Then
                                For Y = 1 To UBound(.AdvMoves)
                                    If .RBYMoves(X) = .AdvMoves(Y) Then .AdvMoves(Y) = .AdvMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.ADVTM)
                                    If .RBYMoves(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvBreeding)
                                    If .RBYMoves(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvSpecial)
                                    If .RBYMoves(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvTutor)
                                    If .RBYMoves(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.RBYTM)
                        If .RBYTM(X) > 0 Then
                            If Not SkipGSC Then
                                For Y = 1 To UBound(.BaseMoves)
                                    If .RBYTM(X) = .BaseMoves(Y) Then .BaseMoves(Y) = .BaseMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MachineMoves)
                                    If .RBYTM(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.BreedingMoves)
                                    If .RBYTM(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.SpecialMoves)
                                    If .RBYTM(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MoveTutor)
                                    If .RBYTM(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                                Next
                            End If
                            If Not SkipAdv Then
                                For Y = 1 To UBound(.AdvMoves)
                                    If .RBYTM(X) = .AdvMoves(Y) Then .AdvMoves(Y) = .AdvMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.ADVTM)
                                    If .RBYTM(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvBreeding)
                                    If .RBYTM(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvSpecial)
                                    If .RBYTM(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvTutor)
                                    If .RBYTM(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                Case 2
                    'Process GSC
                    For X = 1 To UBound(.BaseMoves)
                        If .BaseMoves(X) > 0 Then
                            For Y = 1 To UBound(.MachineMoves)
                                If .BaseMoves(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                            Next
                            For Y = 1 To UBound(.BreedingMoves)
                                If .BaseMoves(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                            Next
                            For Y = 1 To UBound(.SpecialMoves)
                                If .BaseMoves(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                            Next
                            For Y = 1 To UBound(.MoveTutor)
                                If .BaseMoves(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .BaseMoves(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .BaseMoves(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipAdv Then
                                For Y = 1 To UBound(.AdvMoves)
                                    If .BaseMoves(X) = .AdvMoves(Y) Then .AdvMoves(Y) = .AdvMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.ADVTM)
                                    If .BaseMoves(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvBreeding)
                                    If .BaseMoves(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvSpecial)
                                    If .BaseMoves(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvTutor)
                                    If .BaseMoves(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.MachineMoves)
                        If .MachineMoves(X) > 0 Then
                            For Y = 1 To UBound(.BreedingMoves)
                                If .MachineMoves(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                            Next
                            For Y = 1 To UBound(.SpecialMoves)
                                If .MachineMoves(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                            Next
                            For Y = 1 To UBound(.MoveTutor)
                                If .MachineMoves(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .MachineMoves(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .MachineMoves(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipAdv Then
                                For Y = 1 To UBound(.AdvMoves)
                                    If .MachineMoves(X) = .AdvMoves(Y) Then .AdvMoves(Y) = .AdvMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.ADVTM)
                                    If .MachineMoves(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvBreeding)
                                    If .MachineMoves(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvSpecial)
                                    If .MachineMoves(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvTutor)
                                    If .MachineMoves(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.BreedingMoves)
                        If .BreedingMoves(X) > 0 Then
                            For Y = 1 To UBound(.SpecialMoves)
                                If .BreedingMoves(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                            Next
                            For Y = 1 To UBound(.MoveTutor)
                                If .BreedingMoves(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .BreedingMoves(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .BreedingMoves(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipAdv Then
                                For Y = 1 To UBound(.AdvMoves)
                                    If .BreedingMoves(X) = .AdvMoves(Y) Then .AdvMoves(Y) = .AdvMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.ADVTM)
                                    If .BreedingMoves(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvBreeding)
                                    If .BreedingMoves(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvSpecial)
                                    If .BreedingMoves(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvTutor)
                                    If .BreedingMoves(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.SpecialMoves)
                        If .SpecialMoves(X) > 0 Then
                            For Y = 1 To UBound(.MoveTutor)
                                If .SpecialMoves(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .SpecialMoves(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .SpecialMoves(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipAdv Then
                                For Y = 1 To UBound(.AdvMoves)
                                    If .SpecialMoves(X) = .AdvMoves(Y) Then .AdvMoves(Y) = .AdvMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.ADVTM)
                                    If .SpecialMoves(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvBreeding)
                                    If .SpecialMoves(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvSpecial)
                                    If .SpecialMoves(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvTutor)
                                    If .SpecialMoves(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.MoveTutor)
                        If .MoveTutor(X) > 0 Then
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .MoveTutor(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .MoveTutor(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipAdv Then
                                For Y = 1 To UBound(.AdvMoves)
                                    If .MoveTutor(X) = .AdvMoves(Y) Then .AdvMoves(Y) = .AdvMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.ADVTM)
                                    If .MoveTutor(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvBreeding)
                                    If .MoveTutor(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvSpecial)
                                    If .MoveTutor(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                                Next
                                For Y = 1 To UBound(.AdvTutor)
                                    If .MoveTutor(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                Case 3
                    For X = 1 To UBound(.AdvMoves)
                        If .AdvMoves(X) > 0 Then
                            For Y = 1 To UBound(.ADVTM)
                                If .AdvMoves(X) = .ADVTM(Y) Then .ADVTM(Y) = .ADVTM(Y) * -1
                            Next
                            For Y = 1 To UBound(.AdvBreeding)
                                If .AdvMoves(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                            Next
                            For Y = 1 To UBound(.AdvSpecial)
                                If .AdvMoves(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                            Next
                            For Y = 1 To UBound(.AdvTutor)
                                If .AdvMoves(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .AdvMoves(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .AdvMoves(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipGSC Then
                                For Y = 1 To UBound(.BaseMoves)
                                    If .AdvMoves(X) = .BaseMoves(Y) Then .BaseMoves(Y) = .BaseMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MachineMoves)
                                    If .AdvMoves(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.BreedingMoves)
                                    If .AdvMoves(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.SpecialMoves)
                                    If .AdvMoves(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MoveTutor)
                                    If .AdvMoves(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                        For X = 1 To UBound(.ADVTM)
                        If .ADVTM(X) > 0 Then
                            For Y = 1 To UBound(.AdvBreeding)
                                If .ADVTM(X) = .AdvBreeding(Y) Then .AdvBreeding(Y) = .AdvBreeding(Y) * -1
                            Next
                            For Y = 1 To UBound(.AdvSpecial)
                                If .ADVTM(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                            Next
                            For Y = 1 To UBound(.AdvTutor)
                                If .ADVTM(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .ADVTM(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .ADVTM(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipGSC Then
                                For Y = 1 To UBound(.BaseMoves)
                                    If .ADVTM(X) = .BaseMoves(Y) Then .BaseMoves(Y) = .BaseMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MachineMoves)
                                    If .ADVTM(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.BreedingMoves)
                                    If .ADVTM(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.SpecialMoves)
                                    If .ADVTM(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MoveTutor)
                                    If .ADVTM(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.AdvBreeding)
                        If .AdvBreeding(X) > 0 Then
                            For Y = 1 To UBound(.AdvSpecial)
                                If .AdvBreeding(X) = .AdvSpecial(Y) Then .AdvSpecial(Y) = .AdvSpecial(Y) * -1
                            Next
                            For Y = 1 To UBound(.AdvTutor)
                                If .AdvBreeding(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .AdvBreeding(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .AdvBreeding(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipGSC Then
                                For Y = 1 To UBound(.BaseMoves)
                                    If .AdvBreeding(X) = .BaseMoves(Y) Then .BaseMoves(Y) = .BaseMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MachineMoves)
                                    If .AdvBreeding(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.BreedingMoves)
                                    If .AdvBreeding(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.SpecialMoves)
                                    If .AdvBreeding(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MoveTutor)
                                    If .AdvBreeding(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.AdvSpecial)
                        If .AdvSpecial(X) > 0 Then
                            For Y = 1 To UBound(.AdvTutor)
                                If .AdvSpecial(X) = .AdvTutor(Y) Then .AdvTutor(Y) = .AdvTutor(Y) * -1
                            Next
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .AdvSpecial(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .AdvSpecial(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipGSC Then
                                For Y = 1 To UBound(.BaseMoves)
                                    If .AdvSpecial(X) = .BaseMoves(Y) Then .BaseMoves(Y) = .BaseMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MachineMoves)
                                    If .AdvSpecial(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.BreedingMoves)
                                    If .AdvSpecial(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.SpecialMoves)
                                    If .AdvSpecial(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MoveTutor)
                                    If .AdvSpecial(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
                    For X = 1 To UBound(.AdvTutor)
                        If .AdvTutor(X) > 0 Then
                            If Not SkipRBY Then
                                For Y = 1 To UBound(.RBYMoves)
                                    If .AdvTutor(X) = .RBYMoves(Y) Then .RBYMoves(Y) = .RBYMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.RBYTM)
                                    If .AdvTutor(X) = .RBYTM(Y) Then .RBYTM(Y) = .RBYTM(Y) * -1
                                Next
                            End If
                            If Not SkipGSC Then
                                For Y = 1 To UBound(.BaseMoves)
                                    If .AdvTutor(X) = .BaseMoves(Y) Then .BaseMoves(Y) = .BaseMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MachineMoves)
                                    If .AdvTutor(X) = .MachineMoves(Y) Then .MachineMoves(Y) = .MachineMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.BreedingMoves)
                                    If .AdvTutor(X) = .BreedingMoves(Y) Then .BreedingMoves(Y) = .BreedingMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.SpecialMoves)
                                    If .AdvTutor(X) = .SpecialMoves(Y) Then .SpecialMoves(Y) = .SpecialMoves(Y) * -1
                                Next
                                For Y = 1 To UBound(.MoveTutor)
                                    If .AdvTutor(X) = .MoveTutor(Y) Then .MoveTutor(Y) = .MoveTutor(Y) * -1
                                Next
                            End If
                        End If
                    Next
            End Select
        Next
    End With
End Sub

Public Function TrueGSCCheck(ByVal PkNum As Integer, ByVal MoveNum As Integer) As Boolean
    Dim IsMove As Boolean
    Dim X As Integer
     
    IsMove = False
    For X = 1 To UBound(BasePKMN(PkNum).BaseMoves)
        If Abs(BasePKMN(PkNum).BaseMoves(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).MachineMoves)
        If Abs(BasePKMN(PkNum).MachineMoves(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).BreedingMoves)
        If Abs(BasePKMN(PkNum).BreedingMoves(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).SpecialMoves)
        If Abs(BasePKMN(PkNum).SpecialMoves(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).MoveTutor)
        If Abs(BasePKMN(PkNum).MoveTutor(X)) = MoveNum Then IsMove = True
    Next
    If PkNum = 83 And MoveNum = 12 Then IsMove = True
    If PkNum = 207 And MoveNum = 55 Then IsMove = True
    If (PkNum = 147 Or PkNum = 146 Or PkNum = 149) And (MoveNum = 61 Or MoveNum = 240) Then IsMove = True
    If (PkNum = 239 Or PkNum = 125 _
        Or PkNum = 172 Or PkNum = 25 Or PkNum = 26 _
        Or PkNum = 173 Or PkNum = 35 Or PkNum = 36 _
        Or PkNum = 174 Or PkNum = 39 Or PkNum = 40 _
        Or PkNum = 238 Or PkNum = 124 _
        Or PkNum = 240 Or PkNum = 126 _
        Or PkNum = 236 Or PkNum = 106 Or PkNum = 107 Or PkNum = 237) _
        And MoveNum = 45 Then IsMove = True
    TrueGSCCheck = IsMove
End Function
Public Function TrueAdvCheck(ByVal PkNum As Integer, ByVal MoveNum As Integer) As Boolean
    Dim IsMove As Boolean
    Dim X As Integer
     
    IsMove = False
    For X = 1 To UBound(BasePKMN(PkNum).AdvMoves)
        If Abs(BasePKMN(PkNum).AdvMoves(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).ADVTM)
        If Abs(BasePKMN(PkNum).ADVTM(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).AdvBreeding)
        If Abs(BasePKMN(PkNum).AdvBreeding(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).AdvSpecial)
        If Abs(BasePKMN(PkNum).AdvSpecial(X)) = MoveNum Then IsMove = True
    Next
    TrueAdvCheck = IsMove
End Function

Public Function TrueAdvPlusCheck(ByVal PkNum As Integer, ByVal MoveNum As Integer) As Boolean
    Dim IsMove As Boolean
    Dim X As Integer
     
    IsMove = False
    For X = 1 To UBound(BasePKMN(PkNum).AdvMoves)
        If Abs(BasePKMN(PkNum).AdvMoves(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).ADVTM)
        If Abs(BasePKMN(PkNum).ADVTM(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).AdvBreeding)
        If Abs(BasePKMN(PkNum).AdvBreeding(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).AdvSpecial)
        If Abs(BasePKMN(PkNum).AdvSpecial(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).AdvTutor)
        If Abs(BasePKMN(PkNum).AdvTutor(X)) = MoveNum Then IsMove = True
    Next
    For X = 1 To UBound(BasePKMN(PkNum).LFOnly)
        If Abs(BasePKMN(PkNum).LFOnly(X)) = MoveNum Then IsMove = True
    Next
    TrueAdvPlusCheck = IsMove
End Function

Public Sub PlaySound(ByVal SoundNum As Integer)
    Dim FileName As String
    Dim Worked As Variant
    
    '>>> Call WriteDebugLog("Sound " & SoundFile(SoundNum) & " requested")
    On Error Resume Next
    Call StopSound
    If SoundNum < LBound(SoundFile) Or SoundNum > UBound(SoundFile) Then Exit Sub
    FileName = SoundFile(SoundNum)
    If Not SoundEnable(SoundNum) Then Exit Sub
    If Not FileExists(SoundFile(SoundNum)) Then Exit Sub
    Select Case LCase(Right(FileName, 3))
        Case "wav", "mid", "mp3", "wma"
            Worked = CloseMultimedia("SFX")
            Worked = OpenMultimedia(MainContainer.hWnd, "SFX", FileName, "MPEGVideo")
            Worked = PlayMultimedia("SFX", vbNullString, vbNullString)
        Case "mod", ".xm", "s3m", ".it"
            With MainContainer.ModSFX
                .DeviceToUse = 1
                .InitBASS MainContainer.hWnd, 44100, False, False
                .BufferLenInSeconds = 1#
                .CloseModule
                .OpenModule FileName, False
                .PlayModule
            End With
    End Select
    '>>> Call WriteDebugLog("Playing started.")
End Sub

Public Sub PlayMusic(ByVal SoundNum As Integer, Optional ByVal LoopIt As Boolean = False)
    Dim FileName As String
    Dim Worked As Variant
    
    '>>> Call WriteDebugLog("Music " & SoundFile(SoundNum) & " requested.")
    On Error Resume Next
    Call StopMusic
    If SoundNum < LBound(SoundFile) Or SoundNum > UBound(SoundFile) Then Exit Sub
    If Not SoundEnable(SoundNum) Then Exit Sub
    If Not FileExists(SoundFile(SoundNum)) Then Exit Sub
    FileName = SoundFile(SoundNum)
    Select Case LCase(Right(FileName, 3))
        Case "wav", "mid", "mp3", "wma"
            Worked = CloseMultimedia("Music")
            Worked = OpenMultimedia(MainContainer.hWnd, "Music", FileName, "MPEGVideo")
            Worked = PlayMultimedia("Music", vbNullString, vbNullString)
            Worked = SetAutoRepeat(MainContainer.hWnd, "Music", vbNullString, vbNullString, LoopIt)
        Case "mod", ".xm", "s3m", ".it"
            With MainContainer.ModPlay
                .DeviceToUse = 1
                .InitBASS MainContainer.hWnd, 44100, False, False
                .BufferLenInSeconds = 1#
                .CloseModule
                .OpenModule FileName, LoopIt
                .PlayModule
            End With
    End Select
    '>>> Call WriteDebugLog("Music started.")
End Sub

Public Sub StopSound()
    Dim Worked As Variant
    
    Worked = CloseMultimedia("SFX")
    MainContainer.ModSFX.CloseModule
    '>>> Call WriteDebugLog("Sound stopped.")
End Sub

Public Sub StopMusic()
    Dim Worked As Variant
    
    Worked = CloseMultimedia("Music")
    MainContainer.ModPlay.CloseModule
    '>>> Call WriteDebugLog("Music stopped.")
End Sub

Private Function ShortenIP(ByVal IP As String) As String
    Dim Temp As String
    Dim Temp2 As String
    Dim X As Integer
    Temp = ""
    Temp2 = ""
    For X = 1 To Len(IP)
        If Mid(IP, X, 1) = "." Then
            Temp2 = Temp2 & FixedHex(Temp, 2)
            Temp = ""
        Else
            Temp = Temp & Mid(IP, X, 1)
        End If
    Next X
    ShortenIP = Temp2 & FixedHex(Temp, 2)
End Function
Public Function GetNextRnd() As Byte
    If Not UseTrueRnd Then
        GetNextRnd = CByte(Int(Rnd * 256))
    ElseIf RndCache = 0 Then
        If RndState <> rQuerying Then RndState = rEmpty
        GetNextRnd = CByte(Int(Rnd * 256))
    Else
        RndCache = RndCache - 1
        If RndCache <= RndThresh And RndState <> rQuerying Then RndState = rEmpty
        GetNextRnd = RndByte(RndCache)
    End If
End Function
Public Function GetSmallRnd(ByVal High As Byte, Optional ByVal Low As Byte = 0) As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    
    If Low > High Then 'Good chance to practice my nifty variable swap method
        High = High Xor Low
        Low = High Xor Low
        High = High Xor Low
    End If
    
    Y = High - Low + 1
    If Not UseTrueRnd Then
        GetSmallRnd = Int(Rnd * Y) + Low
        Exit Function
    End If
    
    '2 and 4 are even breaks, there's no need to waste an entire byte on them
    Select Case Y
    Case 2, 4
        If BitPos = 0 Or (Y = 4 And BitPos = 1) Then
            Z = GetNextRnd
            For X = 8 To 1 Step -1
                If Z - 2 ^ (X - 1) >= 0 Then
                    RndBit(X) = 1
                    Z = Z - 2 ^ (X - 1)
                Else
                    RndBit(X) = 0
                End If
            Next X
            BitPos = 8
        End If
        Z = 0
        For X = 1 To Y / 2
           Z = Z + RndBit(BitPos) * X
           BitPos = BitPos - 1
        Next X
        GetSmallRnd = Z + Low
    Case Else
        Z = 256 \ Y
        Y = Y * Z
        Do
            X = GetNextRnd
        Loop Until X < Y
        GetSmallRnd = (X \ Z) + Low
    End Select
End Function
Public Function GetMoveNum(ByVal MoveName As String) As Integer
    MoveName = Replace$(LCase$(MoveName), " ", "")
    Dim X As Integer
    For X = 1 To UBound(Moves)
        If Replace$(LCase$(Moves(X).Name), " ", "") = MoveName Then
            GetMoveNum = X
            Exit Function
        End If
    Next X
End Function
Public Function GetPokeNum(ByVal iName As String) As Long
    Dim X As Integer
    iName = LCase$(iName)
    For X = 1 To UBound(BasePKMN)
        If LCase$(BasePKMN(X).Name) = iName Then GetPokeNum = X: Exit Function
    Next X
End Function
Public Function GetMoveSpecial(SpecialNum As Integer) As Long
    Dim X As Integer
    For X = 1 To UBound(Moves)
        If Moves(X).SpecialEffect = SpecialNum Then GetMoveSpecial = X: Exit Function
    Next X
End Function
Public Function GetItemNum(iName As String) As Long
    Dim X As Integer
    iName = LCase$(iName)
    For X = 1 To UBound(Item)
        If LCase$(Item(X)) = iName Then GetItemNum = X: Exit Function
    Next X
End Function
Public Function GetTraitNum(ByVal iName As String) As Long
    Dim X As Integer
    iName = LCase$(iName)
    For X = 0 To UBound(AttributeText)
        If LCase$(AttributeText(X)) = iName Then GetTraitNum = X: Exit Function
    Next X
End Function


Public Sub CenterWindow(ByRef iForm As Form, Optional XOffset As Integer = 0, Optional YOffset As Integer = 0)
    On Error GoTo Failed
    If Not iForm.MDIChild Then Exit Sub
    If MainContainer.Width <= 0 Or MainContainer.Height <= 0 Or MainContainer.WindowState = vbMinimized Then Exit Sub
    iForm.Left = (MainContainer.Width - iForm.Width) \ 2 - 75 + XOffset
    iForm.Top = (MainContainer.Height - iForm.Height) \ 2 - 375 + YOffset
    'The Top must be offset due to titlebars
    'The Left must be offset due to... i don't know =\
    Exit Sub
Failed:
    '>>> Call WriteDebugLog("Error: Something went wrong centering a window!")
End Sub

Public Function ConvertMove(OriginalMove As Move, ByVal GameVersion As BattleModes) As Move
    Dim BuildMove As Move
    BuildMove = OriginalMove
    If GameVersion < nbAdvBattle Then
        With BuildMove
            Select Case .ID
            Case 48 'Double-Edge takes 1/4 recoil instead of 1/3.
                .SpecialEffect = 35
            Case 83 'Growth is a Grass type
                .Type = 5
            Case 111 'Lovely Kiss has 10 PP
                .PP = 10
            Case 112 'Low Kick is completely different
                .Accuracy = 90
                .Power = 50
                .SpecialEffect = 16
            Case 116 'Metitate has 30 PP
                .PP = 30
            Case 160 'Razor Wind has 75% Acc.
                .Accuracy = 75
            End Select
            
            'King's Rock differences
            '(Note to self: Consider making this its own DB entry)
            Select Case .ID
            Case 1, 80, 105, 117, 148, 154, 197, 215, 228, 239, 241
                .KingsRock = True
            Case 16, 26, 51, 55, 60, 66, 85, 95, 112, 170, 180, 206, 240, 246, 250
                .KingsRock = False
            End Select
            
        End With
    End If

    If GameVersion < nbGSCBattle Then
        With BuildMove
            Select Case .ID
            Case 17 'Bite is a Normal attack, and only activates 10% of the time
                .Type = 1
                .SpecialPercent = 10
            Case 18 'Blizzard's Acc. is 90
                .Accuracy = 90
            Case 43, 48 'Both Dig's and Double-Egde's Power is 100
                .Power = 100
            Case 45, 167, 237 'Dizzy Punch, Rock Slide, and Tri Attack have no effect
                .SpecialEffect = 0
            Case 60 'Explosion's Power is 340
                .Power = 340
            Case 64 'Fire Blast burns 30% of the time
                .SpecialPercent = 30
            Case 85, 103, 175 'Gust, Karate Chop, and Sand Attack's type is Normal
                .Type = 1
'            Case 90 'Hi Jump Kick's Acc. is 90
'                .Accuracy = 90
            Case 111 'Lovely Kiss has 10 PP
                .PP = 10
            Case 144 'Poison Sting poisons 20% of the time
                .SpecialPercent = 20
            Case 152 'Psychic drops Special 30% of the time
                .SpecialPercent = 30
            Case 159 'Razor Leaf's Power is 55
                .Power = 55
            Case 169 'Rock Throw's Acc. is 65
                .Accuracy = 65
            Case 180 'Self Destruct's Power is 260
                .Power = 260
            Case 230 'Thunder Paralyses only 10% of the time
                .SpecialPercent = 10
            Case 247 'Whirlwind's Acc. is 85
                .Accuracy = 85
            Case 248 'Wing Attack's Power is 35
                .Power = 35
            End Select
        End With
    End If
    ConvertMove = BuildMove
End Function

Private Function BreedCheck(PokeNum As Integer, MoveArray() As Integer, Version As CompatModes) As Boolean
    Dim Temp As String
    With BasePKMN(PokeNum)
        If CompatVersion(Version) = nbGSCBattle Then
            Temp = .BreedIllegals(0)
            If Version = nbTrueGSC Then Temp = Temp & .BreedIllegals(1)
        Else
            Temp = .BreedIllegals(2)
            If Version = nbTrueGSC Then Temp = Temp & .BreedIllegals(3)
        End If
    End With
    BreedCheck = DBIllegalCheck(MoveArray(), Temp)
End Function
Public Function DBIllegalCheck(MoveArray() As Integer, IllegalList As String) As Boolean
    Dim X As Long
    Dim B As Boolean
    Dim Sets() As String
    Dim TempSet() As String
    B = False
    Sets = Split(IllegalList, "|")
    For X = 1 To UBound(Sets)
        TempSet = Split(Sets(X), "+")
        Select Case UBound(TempSet)
        Case 0: B = HasMoves(MoveArray(), TempSet(0))
        Case 1: B = HasMoves(MoveArray(), TempSet(0), TempSet(1))
        Case 2: B = HasMoves(MoveArray(), TempSet(0), TempSet(1), TempSet(2))
        Case 3: B = HasMoves(MoveArray(), TempSet(0), TempSet(1), TempSet(2), TempSet(3))
        End Select
        If B Then Exit For
    Next X
    DBIllegalCheck = Not B
End Function
Public Function HasMoves(MoveArray() As Integer, ByVal M1 As Integer, Optional ByVal M2 As Integer = 0, Optional ByVal M3 As Integer = 0, Optional ByVal M4 As Integer = 0) As Boolean
    Dim Check() As Integer
    Dim A As Integer
    Dim X As Integer
    Dim Y As Integer
    If M4 = 0 Then
        If M3 = 0 Then
            If M2 = 0 Then
                ReDim Check(1 To 1)
            Else
                ReDim Check(1 To 2)
            End If
        Else
            ReDim Check(1 To 3)
        End If
    Else
        ReDim Check(1 To 4)
    End If
    On Error Resume Next
    Check(1) = M1
    Check(2) = M2
    Check(3) = M3
    Check(4) = M4
    A = 0
    For X = 1 To UBound(Check)
        For Y = LBound(MoveArray) To UBound(MoveArray)
            If MoveArray(Y) <> 0 And MoveArray(Y) = Check(X) Then
                A = A + 1
                Exit For
            End If
        Next Y
    Next X
    HasMoves = (A = UBound(Check))
End Function


Public Function OtherTeam(ByVal Team As Long) As Byte
    If Team = 1 Then OtherTeam = 2 Else OtherTeam = 1
End Function

Public Function Cap(ByVal Value As Long, Optional ByVal CapNum As Long = 999) As Long
    If Value > CapNum Then Cap = CapNum Else Cap = Value
End Function
Public Function Minimum(ByVal Value As Long, Optional ByVal MinVal As Long = 0) As Long
    If Value < MinVal Then Minimum = MinVal Else Minimum = Value
End Function
Public Sub SetEnglish()
    Dim X As Integer
    Dim Y As Long
    Dim Temp As String
    Dim Build As String
    
    '>>> Call WriteDebugLog("Setting English strings...")
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
    
    'Descriptions for the genders
    Gender(0) = "Genderless"
    Gender(1) = "Male"
    Gender(2) = "Female"
    
    'Descriptions for the conditions
    Condition(1) = "Nrm"
    Condition(2) = "Psn"
    'Note: 3 is actually Toxic, it remins separate due to Baton Pass weirdness.
    Condition(3) = "Psn"
    Condition(4) = "Slp"
    Condition(5) = "Brn"
    Condition(6) = "Par"
    Condition(7) = "Frz"
    Condition(8) = "Fnt"

    'Item names
    Item(0) = "(No Item)"
    Item(1) = "Berry"
    Item(2) = "Berry Juice"
    Item(3) = "Bitter Berry"
    Item(4) = "Burnt Berry"
    Item(5) = "Gold Berry"
    Item(6) = "Ice Berry"
    Item(7) = "Mint Berry"
    Item(8) = "Miracle Berry"
    Item(9) = "Mysteryberry"
    Item(10) = "PrzCureBerry"
    Item(11) = "PsnCureBerry"
    Item(12) = "Berserk Gene"
    Item(13) = "Black Belt"
    Item(14) = "BlackGlasses"
    Item(15) = "BrightPowder"
    Item(16) = "Charcoal"
    Item(17) = "Dragon Fang"
    Item(18) = "Focus Band"
    Item(19) = "Hard Stone"
    Item(20) = "King's Rock"
    Item(21) = "Leftovers"
    Item(22) = "Light Ball"
    Item(23) = "Lucky Punch"
    Item(24) = "Magnet"
    Item(25) = "Metal Coat"
    Item(26) = "Metal Powder"
    Item(27) = "Miracle Seed"
    Item(28) = "Mystic Water"
    Item(29) = "NevermeltIce"
    Item(30) = "Pink Bow"
    Item(31) = "Poison Barb"
    Item(32) = "Polkadot Bow"
    Item(33) = "Quick Claw"
    Item(34) = "Scope Lens"
    Item(35) = "Sharp Beak"
    Item(36) = "Silver Powder"
    Item(37) = "Soft Sand"
    Item(38) = "Spell Tag"
    Item(39) = "Stick"
    Item(40) = "Thick Club"
    Item(41) = "TwistedSpoon"
    Item(42) = "Cheri Berry"
    Item(43) = "Chesto Berry"
    Item(44) = "Pecha Berry"
    Item(45) = "Rawst Berry"
    Item(46) = "Aspear Berry"
    Item(47) = "Leppa Berry"
    Item(48) = "Oran Berry"
    Item(49) = "Persim Berry"
    Item(50) = "Lum Berry"
    Item(51) = "Sitrus Berry"
    Item(52) = "Figy Berry"
    Item(53) = "Iapapa Berry"
    Item(54) = "Mago Berry"
    Item(55) = "Wiki Berry"
    Item(56) = "Aguav Berry"
    Item(57) = "Liechi Berry"
    Item(58) = "Ganlon Berry"
    Item(59) = "Salac Berry"
    Item(60) = "Petaya Berry"
    Item(61) = "Apicot Berry"
    Item(62) = "Lansat Berry"
    Item(63) = "Starf Berry"
    Item(64) = "Choice Band"
    Item(65) = "DeepSeaScale"
    Item(66) = "DeepSeaTooth"
    Item(67) = "Lax Incense"
    Item(68) = "Macho Brace"
    Item(69) = "Mental Herb"
    Item(70) = "Sea Incense"
    Item(71) = "Shell Bell"
    Item(72) = "Silk Scarf"
    Item(73) = "Soul Dew"
    Item(74) = "White Herb"
    
    ItemDesc(1) = "A Berry that restores 10 HP when the holding Pokémon falls below half its total HP."
    ItemDesc(2) = "Juice that restores 20 HP when the holding Pokémon falls below half its total HP."
    ItemDesc(3) = "A Berry that cures Confusion.  It activates as soon as the holding Pokémon becomes Confused."
    ItemDesc(4) = "A Berry that cures Freeze.  It activates as soon as the holding Pokémon becomes Frozen."
    ItemDesc(5) = "A Berry that restores 30 HP when the holding Pokémon falls below half its total HP."
    ItemDesc(6) = "A Berry that cures Burn.  It activates as soon as the holding Pokémon becomes Burnt."
    ItemDesc(7) = "A Berry that cures Sleep.  It activates as soon as the holding Pokémon falls Asleep."
    ItemDesc(8) = "A Berry that cures all status ailments.  It activates as soon as the holding Pokémon receives a status ailment."
    ItemDesc(9) = "A Berry that restores 5 PP when the PP of any of the Pokémon's moves falls to zero."
    ItemDesc(10) = "A Berry that cures Paralysis.  It activates as soon as the holding Pokémon falls Paralyzed."
    ItemDesc(11) = "A Berry that cures Poison.  It activates as soon as the holding Pokémon falls Poisoned."
    ItemDesc(12) = "A Gene that sharply raises Attack but confuses the holding Pokémon.  It activates as soon as the holding Pokémon enters play."
    ItemDesc(13) = "A cloth belt that raises the power of the holding Pokémon's Fighting type moves by 10%."
    ItemDesc(14) = "A pair of glasses that raises the power of the holding Pokémon's Dark type moves by 10%."
    ItemDesc(15) = "Shining powder that slightly increases the holding Pokémon's evasiveness."
    ItemDesc(16) = "A piece of charcoal that raises the power of the holding Pokémon's Fire type moves by 10%."
    ItemDesc(17) = "A long tooth that raises the power of the holding Pokémon's Dragon type moves by 10%."
    ItemDesc(18) = "A headband that gives the holding Pokémon a 12% chance of enduring a lethal attack with 1 HP."
    ItemDesc(19) = "A smooth rock that raises the power of the holding Pokémon's Rock type moves by 10%."
    ItemDesc(20) = "A rock with a strange symbol that gives certain attacks a 12% chance of causing the target to Flinch."
    ItemDesc(21) = "Leftover food that restores 1/16 of the holding Pokémon's total HP every turn."
    ItemDesc(22) = "A glowing orb that doubles the Special Attack of Pikachu."
    ItemDesc(23) = "A glove that gives Chansey a high Critical Hit rate."
    ItemDesc(24) = "A horseshoe magnet that raises the power of the holding Pokémon's Electric type moves by 10%."
    ItemDesc(25) = "A metallic coating that raises the power of the holding Pokémon's Steel type moves by 10%."
    ItemDesc(26) = "A powder that raises Ditto's Defense and Special Defense by 1.5x."
    ItemDesc(27) = "A seed that raises the power of the holding Pokémon's Grass type moves by 10%."
    ItemDesc(28) = "A pouch of water that raises the power of the holding Pokémon's Water type moves by 10%."
    ItemDesc(29) = "A small block of ice that raises the power of the holding Pokémon's Ice type moves by 10%."
    ItemDesc(30) = "A pretty bow that raises the power of the holding Pokémon's Normal type moves by 10%."
    ItemDesc(31) = "A thorn that raises the power of the holding Pokémon's Poison type moves by 10%."
    ItemDesc(32) = "A pretty bow that raises the power of the holding Pokémon's Normal type moves by 10%."
    ItemDesc(33) = "A talon that has a 24% chance of allowing the holding Pokémon to strike first."
    ItemDesc(34) = "A lens that raises the holding Pokémon's Critical Hit rate."
    ItemDesc(35) = "A sharpened beak that raises the power of the holding Pokémon's Flying type moves by 10%."
    ItemDesc(36) = "A silvery powder that raises the power of the holding Pokémon's Bug type moves by 10%."
    ItemDesc(37) = "Sand that raises the power of the holding Pokémon's Ground type moves by 10%."
    ItemDesc(38) = "A mysterious tag that raises the power of the holding Pokémon's Ghost type moves by 10%."
    ItemDesc(39) = "A large stick that raises Farfetch'd's Critical Hit rate."
    ItemDesc(40) = "A large bone that doubles Cubone's and Marowak's Attack."
    ItemDesc(41) = "A bent spoon that raises the power of the holding Pokémon's Psychic type moves by 10%."
    ItemDesc(42) = "A Berry that cures Paralysis.  It activates as soon as the holding Pokémon falls Paralyzed."
    ItemDesc(43) = "A Berry that cures Sleep.  It activates as soon as the holding Pokémon falls Asleep."
    ItemDesc(44) = "A Berry that cures Poison.  It activates as soon as the holding Pokémon falls Poisoned."
    ItemDesc(45) = "A Berry that cures Burn.  It activates as soon as the holding Pokémon becomes Burned."
    ItemDesc(46) = "A Berry that cures Freeze.  It activates as soon as the holding Pokémon becomes Frozen."
    ItemDesc(47) = "A Berry that restores 10 PP at the end of the turn if the PP of any of the Pokémon's moves is at zero."
    ItemDesc(48) = "A Berry that restores 10 HP when the holding Pokémon falls below half its total HP."
    ItemDesc(49) = "A Berry that cures Confusion.  It activates as soon as the holding Pokémon becomes Confused."
    ItemDesc(50) = "A Berry that cures all status ailments.  It activates as soon as the holding Pokémon receives a status ailment."
    ItemDesc(51) = "A Berry that restores 30 HP when the holding Pokémon falls below half its total HP."
    ItemDesc(52) = "A Berry that restores 1/8 of the holding Pokémon's total HP when it falls below half.  It may cause Confusion for some Pokémon."
    ItemDesc(53) = "A Berry that restores 1/8 of the holding Pokémon's total HP when it falls below half.  It may cause Confusion for some Pokémon."
    ItemDesc(54) = "A Berry that restores 1/8 of the holding Pokémon's total HP when it falls below half.  It may cause Confusion for some Pokémon."
    ItemDesc(55) = "A Berry that restores 1/8 of the holding Pokémon's total HP when it falls below half.  It may cause Confusion for some Pokémon."
    ItemDesc(56) = "A Berry that restores 1/8 of the holding Pokémon's total HP when it falls below half.  It may cause Confusion for some Pokémon."
    ItemDesc(57) = "A Berry that raises Attack when the holding Pokémon falls below 1/4 its total HP."
    ItemDesc(58) = "A Berry that raises Defense when the holding Pokémon falls below 1/4 its total HP."
    ItemDesc(59) = "A Berry that raises Speed when the holding Pokémon falls below 1/4 its total HP."
    ItemDesc(60) = "A Berry that raises Special Attack when the holding Pokémon falls below 1/4 its total HP."
    ItemDesc(61) = "A Berry that raises Special Defense when the holding Pokémon falls below 1/4 its total HP."
    ItemDesc(62) = "A Berry that raises the Critical Hit rate when the holding Pokémon falls below 1/4 its total HP."
    ItemDesc(63) = "A Berry that sharply raises a random stat when the holding Pokémon falls below 1/4 its total HP."
    ItemDesc(64) = "A headband that raises Attack by 1.5x, disallows the use of all of the holding Pokémon's moves except the first it uses."
    ItemDesc(65) = "A scale that doubles Clamperl's Special Defense."
    ItemDesc(66) = "A scale that doubles Clamperl's Special Attack."
    ItemDesc(67) = "Incense that slightly increases the holding Pokémon's evasiveness."
    ItemDesc(68) = "A brace that doubles the Effort Points received after battles with wild Pokémon.  Its only in-battle effect is that it halves Speed."
    ItemDesc(69) = "An herb that cures Attraction.  It activates as soon as the holding Pokémon falls Attracted."
    ItemDesc(70) = "Incense that raises the power of the holding Pokémon's Water type moves by 10%."
    ItemDesc(71) = "A bell that restores 1/8 of the damage the holding Pokémon inflicts when it attacks."
    ItemDesc(72) = "A pretty scarf that raises the power of the holding Pokémon's Normal type moves by 10%."
    ItemDesc(73) = "A glittering gem that increases the Special Attack and Special Defense of Latias and Latios by 1.5x."
    ItemDesc(74) = "An herb that restores stats.  It activates as soon as any of holding Pokémon's stats are reduced."
    
    'Natures
    Nature(0).Name = "Hardy"
    Nature(1).Name = "Lonely"
    Nature(2).Name = "Brave"
    Nature(3).Name = "Adamant"
    Nature(4).Name = "Naughty"
    Nature(5).Name = "Bold"
    Nature(6).Name = "Docile"
    Nature(7).Name = "Relaxed"
    Nature(8).Name = "Impish"
    Nature(9).Name = "Lax"
    Nature(10).Name = "Timid"
    Nature(11).Name = "Hasty"
    Nature(12).Name = "Serious"
    Nature(13).Name = "Jolly"
    Nature(14).Name = "Naive"
    Nature(15).Name = "Modest"
    Nature(16).Name = "Mild"
    Nature(17).Name = "Quiet"
    Nature(18).Name = "Bashful"
    Nature(19).Name = "Rash"
    Nature(20).Name = "Calm"
    Nature(21).Name = "Gentle"
    Nature(22).Name = "Sassy"
    Nature(23).Name = "Careful"
    Nature(24).Name = "Quirky"

    'Weather Conditions
    Weather(0) = "Normal"
    Weather(1) = "Raining"
    Weather(2) = "Sunny"
    Weather(3) = "Sandstorm"
    Weather(4) = "Hailstorm"
    
    'Rule Text
    RuleText(nbSleep) = "Sleep Clause"
    RuleText(nbFreeze) = "Freeze Clause"
    RuleText(nbSelfKO) = "Self-KO Clause"
    RuleText(nbUsePPUps) = "Apply PP Ups"
    RuleText(nbStadiumMode) = "Stadium Mode"
    RuleText(nbLevelBalance) = "Level Balance"
    RuleText(nbNoWatch) = "Disallow Spectators"
    RuleText(nbTimeout) = "Battle Timeout"
    RuleText(nbUnrated) = "Unrated"
    RuleText(nbPresentRule) = "Use Stadium Present"
    RuleText(nbExactHP) = "Show Enemy HP"
    RuleText(nbRandbat) = "Challenge Cup"

    'No more restrictions! =D
    'Full Rule Text
    RuleToolTip(nbSleep) = "With this rule enabled, only one Pokémon can be Asleep %nper team.  If a move is used that would violate this%nrule, it will automatically fail.  Pokémon under the%neffects of the move Rest do not count for this rule."
    RuleToolTip(nbFreeze) = "With this rule enabled, only one Pokémon can be Frozen %nper team."
    RuleToolTip(nbSelfKO) = "With this rule enabled, if your team has only one Pokémon%nremaining, the moves Destiny Bond and Perish Song will%nalways fail when used by that Pokémon.  Also, if your final%nPokémon uses SelfDestruct or Explosion to Knock Out your%nopponent's final Pokémon, you will automatically lose.  The%npurpose of this rule is to prevent the match from ending in a%nTie."
'    RuleToolTip(nbNo1HKOs) = "With this rule enabled, the moves Fissure, Horn Drill, and%nGuillotine will always fail for both players."
    RuleToolTip(nbUsePPUps) = "With this rule enabled, all moves on both your and your%nopponent's Pokémon will have 60% more Power Points."
    RuleToolTip(nbStadiumMode) = "With this rule enabled, the match will be 3 on 3.  Each player%nwill be allowed to pick three of their six Pokémon to use in%nbattle.  The remaining three Pokémon will not be used."
'    RuleToolTip(nbRBYMode) = "With this rule enabled, the battle mechanics and formulas will%nsimulate those of Pokémon Red/Blue/Yellow.  Some moves%nmay have different powers, accuracies, or effects from those%nin  the standard Gold/Silver/Crystal mode.  This rule cannot be%nselected unless both your and your opponent's teams are%nR/B/Y Compatible."
    RuleToolTip(nbRandbat) = "With this rule enabled, each player will use six random%nPokémon instead of his/her own team.  The moves of each%nPokémon will also be random, however, each one will always%nhave at least one damaging attack."
    RuleToolTip(nbLevelBalance) = "With this rule enabled, the Levels of each Pokémon will be%naltered to balance its power.  Pokémon with unusually high%nstats will be assigned a lower level, and Pokémon with lower%nstats will be assigned a higher level."
    RuleToolTip(nbNoWatch) = "With this rule enabled, spectators will be unable to watch your%nbattle.  However, mods and admins will still be able to watch for%nmoderation purposes."
'    RuleToolTip(nbNoWatchChat) = "With this rule enabled, spectators will be able to watch your%nbattle, but will not be allowed to chat."
    RuleToolTip(nbExactHP) = "With this rule enabled, both players will see the exact HP of%nthe opposing Pokémon instead of a percentage."
    RuleToolTip(nbTimeout) = "With this rule enabled, after one player has chosen a move, a%ntimer will start.  After 5 minutes of inactivity, the battle will%nautomatically end and the active player will be awarded the%nwin."
    RuleToolTip(nbUnrated) = "With this rule enabled, the outcome of the battle will affect%nneither players' battle record."
    RuleToolTip(nbPresentRule) = "With this rule enabled, the formulas used to calculate the%ndamage for the move Present will be those of Pokémon%nCrystal and Pokémon Stadium 2.  If not enabled, the formulas%nin Pokémon Gold and Pokémon Silver will be used."
    

'    RuleToolTip(nbSleepFreeze) = "With this rule enabled, only one Pokémon can be Asleep or  Frozen per team.  If a move is used that would violate this rule, it will automatically fail.  However, it is possible to have one Pokémon Frozen and one Asleep.  Pokémon under the effects of the move Rest do not count for this rule."
'    RuleToolTip(nbSelfKO) = "With this rule enabled, if your team has only one Pokémon remaining, the moves Destiny Bond and Perish Song will always fail when used by that Pokémon.  Also, if your final Pokémon uses Self-Destruct or Explosion to Knock Out your opponent's final Pokémon, you will automatically lose.  The purpose of this rule is to prevent the match from ending in a Tie."
'    RuleToolTip(nbNo1HKOs) = "With this rule enabled, the moves Fissure, Horn Drill, and Guillotine will always fail for both players."
'    RuleToolTip(nbUsePPUps) = "With this rule enabled, all moves on both your and your  opponent's Pokémon will have 60% more Power Points."
'    RuleToolTip(nbIndoorsMode) = "With this rule enabled, the match will be 3 on 3.  Each player  will be allowed to pick three of their six Pokémon to use in battle.  The remaining three Pokémon will not be used."
'    RuleToolTip(nbRBYMode) = "With this rule enabled, the battle mechanics and formulas will simulate those of Pokémon Red/Blue/Yellow.  Some moves may have different powers, accuracies, or effects from those in  the standard Gold/Silver/Crystal mode.  This rule cannot be selected unless both your and your opponent's teams are R/B/Y Compatible."
'    RuleToolTip(nbChallengeCup) = "With this rule enabled, each player will use six random Pokémon instead of his/her own team.  The moves of each Pokémon will also be random, however, each one will always have at least one damaging attack."
'    RuleToolTip(nbLevelBalance) = "With this rule enabled, the Levels of each Pokémon will be altered to balance its power.  Pokémon with usually high stats will be assigned a lower level, and Pokémon with lower stats will be assigned a higher level."
'    RuleToolTip(nbNoWatch) = "With this rule enabled, spectators will be unable to watch your battle."
'    RuleToolTip(nbNoWatchChat) = "With this rule enabled, spectators will be able to watch your battle, but will not be allowed to chat."
'    RuleToolTip(nbTimeout) = "With this rule enabled, after one player has chosen a move, a timer will start.  After 5 minutes of inactivity, the battle will automatically end and the active player will be awarded the win."
'    RuleToolTip(nbUnrated) = "With this rule enabled, the outcome of the battle will affect neither players' battle record."
'    RuleToolTip(nbPresentRule) = "With this rule enabled, the formulas used to calculate the damage for the move Present will be those of Pokémon Crystal and Pokémon Stadium 2.  If not enabled, the formulas in Pokémon Gold and Pokémon Silver will be used."

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
    EvoMethod(9) = "Happiness (Day)"
    EvoMethod(10) = "Happiness (Night)"
    EvoMethod(11) = "Trade (With Item)"
    EvoMethod(12) = "Moon Stone"
    EvoMethod(13) = "Beauty"
    EvoMethod(14) = "Egg"
    EvoMethod(15) = "Egg (With Item)"
    
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
    'Using translation from Pokemon Forever - adjust values when the English version releases.
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
    AttributeText(14) = "Compoundeyes"
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
    
    AttributeDesc(0) = ""
    AttributeDesc(1) = "Pokémon with this Trait posses an unbearable stench that keeps away wild Pokémon.  It has no effect in battle."
    AttributeDesc(2) = "Pokémon with this Trait cause Rain to fall for the remainder of the battle when they enter play."
    AttributeDesc(3) = "The Speed of any Pokémon with this Trait rises by one level at the end of every turn."
    AttributeDesc(4) = "Pokémon with this Trait have a thick layer of armor that prevents Critical Hits."
    AttributeDesc(5) = "Pokémon with this Trait are completely immune to all One-Hit KO attacks."
    AttributeDesc(6) = "When a Pokémon with this Trait is active, the battle arena becomes so damp that no active Pokémon can use Explosion or Selfdestruct."
    AttributeDesc(7) = "Pokémon with this Trait have extremely flexible bodies that are unable to be Paralyzed."
    AttributeDesc(8) = "Pokémon with this Trait are slightly more evasive during a Sandstorm."
    AttributeDesc(9) = "When a Pokémon with this Trait is attacked with a Contact Move, there is a 30% chance of the attacker becoming Paralyzed."
    AttributeDesc(10) = "When a Pokémon with this Trait is attacked with an Electric type move, it will absorb the damage and recover a maximum of 1/4 its total HP."
    AttributeDesc(11) = "When a Pokémon with this Trait is attacked with a Water type move, it will absorb the damage and recover a maximum of 1/4 its total HP."
    AttributeDesc(12) = "Pokémon with this Trait are immune to the move Attract."
    AttributeDesc(13) = "When a Pokémon with this Trait is active, all effects of the current weather are completely negated."
    AttributeDesc(14) = "Pokémon with this Trait have large eyes, making all attacks 30% more accurate."
    AttributeDesc(15) = "Pokémon with this Trait are unable to fall Asleep."
    AttributeDesc(16) = "When a Pokémon with this Trait is attacked, it will change its type to that of the move it was hit with."
    AttributeDesc(17) = "Pokémon with this move have a natural immunity to all forms of Poison."
    AttributeDesc(18) = "When a Pokémon with this Trait is attacked with a Fire type move, it becomes immune to all Fire type moves, and its own Fire type moves become 1.5x more powerful."
    AttributeDesc(19) = "Pokémon with this Trait are coated in a strange powder that the negates the extra effects of most moves."
    AttributeDesc(20) = "Pokémon with this Trait have unique minds and cannot become Confused."
    AttributeDesc(21) = "Pokémon with this Trait use suction cups to anchor themselves, negating the effects of Roar and Whirlwind."
    AttributeDesc(22) = "When Pokémon with this Trait enter play, their presence is so intimidating that any opposing Pokémon has its Attack reduced by one level."
    AttributeDesc(23) = "Pokémon with this Trait possess the ability to freeze the shadows of opponents and prevent them from leaving play."
    AttributeDesc(24) = "Pokémon with this Trait have extremely coarse skin that damages any Pokémon that uses a Contact Move against it by 1/16 of their total HP."
    AttributeDesc(25) = "Pokémon with this Trait are protected by a mystical barrier that negates damage from all attacks except those that are ""Super Effective"" against them."
    AttributeDesc(26) = "Pokémon with this Trait can float in the air, thereby avoiding all Ground type attacks."
    AttributeDesc(27) = "When a Pokémon with this Trait is attacked with a Contact Move, there is a 10% chance of the attacker becoming either Paralyzed, Poisoned, or Asleep."
    AttributeDesc(28) = "When a Pokémon with this Trait becomes Paralyzed, Poisoned, or Burned, the Pokémon that inflicted it receives the same status as well."
    AttributeDesc(29) = "Pokémon with this Trait are completely immune to all stat-lowering attacks."
    AttributeDesc(30) = "Pokémon with this Trait posses the ability to cure themselves upon leaving play."
    AttributeDesc(31) = "When a Pokémon with this Trait is in play, all Electric attacks are drawn to it, regardless of the intended target."
    AttributeDesc(32) = "When a Pokémon with this Trait uses an attack that has an extra effect, the chances of the effect occurring are doubled."
    AttributeDesc(33) = "Pokémon with this Trait use their incredible swimming ability to double their Speed while it is Raining."
    AttributeDesc(34) = "Pokémon with this Trait absorb sunlight to double their Speed while it is Sunny."
    AttributeDesc(35) = "Pokémon with this Trait emit a warm glow that attracts wild Pokémon.  It has no effect in battle."
    AttributeDesc(36) = "When a Pokémon with this Trait enters play, it copies the Trait of an opposing Pokémon until it leaves play."
    AttributeDesc(37) = "Pokémon with this Trait have double their normal Attack."
    AttributeDesc(38) = "When a Pokémon with this Trait is attacked with a Contact Move, there is a 30% chance of the attacker becoming Poisoned."
    AttributeDesc(39) = "Pokémon with this Trait concentrate intensely to prevent Flinching."
    AttributeDesc(40) = "Pokémon with this Trait are covered in a burning exterior that prevents Freezing."
    AttributeDesc(41) = "Pokémon with this Trait are surrounded by a shield of water that prevents Burns."
    AttributeDesc(42) = "Pokémon with this Trait are magnetically charged and will prevent all Steel type Pokémon from leaving play."
    AttributeDesc(43) = "Pokémon with this Trait are completely immune to all Sound Moves."
    AttributeDesc(44) = "Pokémon with this Trait use a dish atop their heads to collect rainwater and use it to heal themselves by 1/16 of their total HP each turn it Rains."
    AttributeDesc(45) = "Pokémon with this Trait cause a ferocious Sandstorm to rage for the remainder of the battle when they enter play."
    AttributeDesc(46) = "When a Pokémon with this Trait is attacked, the attacker loses an extra PP from the attack it used."
    AttributeDesc(47) = "Pokémon with this Trait receive half damage from Fire type and Ice type attacks."
    AttributeDesc(48) = "When a Pokémon with this Trait falls asleep, the duration of its sleep will be only half as long as it would have been."
    AttributeDesc(49) = "When a Pokémon with this Trait is attacked with a Contact Move, there is a 30% chance of the attacker becoming Burned."
    AttributeDesc(50) = "Pokémon with this Trait can flee from wild Pokémon without fail.  It has no effect in battle."
    AttributeDesc(51) = "Pokémon with this Trait have exceptional vision, and their Accuracy cannot be lowered."
    AttributeDesc(52) = "Pokémon with this Trait are extremely fierce, and their Attack cannot be lowered."
    AttributeDesc(53) = "Pokémon with this Trait have a chance to pick up items after battles with wild Pokémon.  It has no effect in battle."
    AttributeDesc(54) = "Pokémon with this Trait are extremely lazy, and cannot attack consecutively."
    AttributeDesc(55) = "Physical attacks from Pokémon with this Trait do 1.5x more damage, but their accuracy is 20% lower."
    AttributeDesc(56) = "When a Pokémon with this Trait is attacked with a Contact Move, there is a 30% chance of the attacker becoming Attracted if the attacker is of the opposite gender."
    AttributeDesc(57) = "If the Ally of a Pokémon with this Trait has the Trait ""Minus,"" its Special Attack will be 1.5x greater."
    AttributeDesc(58) = "If the Ally of a Pokémon with this Trait has the Trait ""Plus,"" its Special Attack will be 1.5x greater."
    AttributeDesc(59) = "Pokémon with this Trait change their type according to the weather."
    AttributeDesc(60) = "Pokémon with this Trait cannot have their item stolen in any way."
    AttributeDesc(61) = "When a Pokémon with this Trait has a status ailment, there is a 30% chance each turn of the Pokémon shedding the ailment."
    AttributeDesc(62) = "When a Pokémon with this Trait receives a status ailment, its Attack is increased by 1.5x."
    AttributeDesc(63) = "When a Pokémon with this Trait receives a status ailment, its Defense is increased by 1.5x."
    AttributeDesc(64) = "When Pokémon with this Trait are attacked by a HP draining attack, such as Giga Drain, the attacker's HP will drop instead of rise."
    AttributeDesc(65) = "When the HP of a Pokémon with this Trait falls below 1/3 it's total HP, that Pokémon's Grass type attacks will do 1.5x damage."
    AttributeDesc(66) = "When the HP of a Pokémon with this Trait falls below 1/3 it's total HP, that Pokémon's Fire type attacks will do 1.5x damage."
    AttributeDesc(67) = "When the HP of a Pokémon with this Trait falls below 1/3 it's total HP, that Pokémon's Water type attacks will do 1.5x damage."
    AttributeDesc(68) = "When the HP of a Pokémon with this Trait falls below 1/3 it's total HP, that Pokémon's Bug type attacks will do 1.5x damage."
    AttributeDesc(69) = "Pokémon with this Trait have a thick skull that prevents recoil damage from moves such as Double-Edge."
    AttributeDesc(70) = "Pokémon with this Trait cause the weather to be Sunny for the remainder of the battle when they enter play."
    AttributeDesc(71) = "Pokémon with this Trait prevent any opposing Pokémon from leaving play, except those that are Flying type or possess the Trait ""Levitate."""
    AttributeDesc(72) = "Pokémon with this Trait are extremely alert and cannot fall Asleep."
    AttributeDesc(73) = "Pokémon with this Trait are protected by a veil of smoke that negates the effects of all stat-lowering moves."
    AttributeDesc(74) = "Pokémon with this Trait have double their normal Attack."
    AttributeDesc(75) = "Pokémon with this Trait have a thick layer of armor that prevents Critical Hits."
    AttributeDesc(76) = "Pokémon with this Trait negate the effects of Sound Moves by drowning them out."
    AttributeDesc(77) = "When a Pokémon with this Trait is active, all effects of the current weather are completely negated."
    
    'Stats
    StatName(1) = "HP"
    StatName(2) = "Attack"
    StatName(3) = "Defense"
    StatName(4) = "Speed"
    StatName(5) = "Special"
    StatName(6) = "Special Attack"
    StatName(7) = "Special Defense"
    StatName(8) = "Accuracy"
    StatName(9) = "Evasion"
    
    
    X = FreeFile
    Open SlashPath & "EnglishText.pnf" For Binary Access Read As #X
    Build = String$(LOF(X), vbNullChar)
    Get #X, , Build
    Close #X
    Temp = ChopString(Build, 2)
    X = Asc(Mid(Temp, 1, 1)) * 256
    X = X + Asc(Mid(Temp, 2, 1))
    Temp = ChopString(Build, X)
    ReDim FTextOffset(X)
    ReDim FTextLen(X)
    Y = 0
    For X = 1 To X
        FTextOffset(X) = Y + 1
        FTextLen(X) = Asc(Mid(Temp, X, 1))
        Y = Y + FTextLen(X)
    Next X
    Call MainContainer.Compressor.DecompressString(Build, Y)
    FTextFile = SlashPath & "TEMPFT" & FileHex & IIf(InVBMode, ".vbtmp", ".tmp")
    Do
        FTextFile = SlashPath & "TEMPFT" & FixedHex(CLng(Rnd * 65535), 4) & IIf(InVBMode, ".vbtmp", ".tmp")
    Loop Until Dir(FTextFile) = ""
    X = FreeFile
    Open FTextFile For Binary Access Write As #X
    Put #X, , Build
    Close #X
    Call SetAttr(FTextFile, vbHidden + vbReadOnly)
    
    'The actual strings are stored elsewhere, but use this for refrence.
    
    '******GENERAL******
    '1 -   %1 (Lv.%2 %3), go!
    '2 -   %1, return!
    '3 -   %1 fainted!
    '4 -   %1 used %2!
    '5 -   A critical hit!
    '6 -   It's a one-hit KO!
    '7 -   It's not very effective...
    '8 -   It's super effective!
    '9 -   Hit!
    '10 -  Hit %1 times!
    '11 -  %1 is hit with recoil!
    '12 -  %1's team's %2 wore off!
    '13 -  %1's %2 rose!
    '14 -  %1's %2 fell!
    '15 -  %1's %2 sharply rose!
    '16 -  %1's %2 sharply fell!
    '17 -  (%1 damage)
    '18 -  (%1% damage)
    '19 -  (%1 damage to the Substitute)
    '20 -  (%1% damage to the Substitute)
    '21 -  (%1 damage to %2)
    '22 -  (%1% damage to %2)
    '23 -  Rule: %1
    '24 -  End of turn #%1
    '25 -  %1's %2: %3 HP
    '26 -  %1's %2: %3% HP
    '27 -  %1's %2: %3 HP (%4)
    '28 -  %1's %2: %3% HP (%4)
    '29 -  %1 won!
    '30 -  %1 lost!
    '31 -  End Battle!  The match is a Tie!
    '32 -  Begin Turn #%1
    '33 -  %1 vs %2.  Begin!
    '34 -  %1 vs %2%nThe battle is already underway.
    '35 -  End Battle!  %1 wins!
    '36 -  Score: %1 to %2
    '
    '******FAILURE MESSAGES******
    '37 -  %1's attack missed!
    '38 -  But it failed!
    '39 -  But nothing happened!
    '40 -  %1 can't move!
    '41 -  %1 protected itself!
    '42 -  %1 avoided damage with %2!
    '43 -  %1 makes %2 moves miss with %3!
    '44 -  %1 avoided the attack!
    '45 -  It doesn't affect %1...
    '46 -  %1 prevents escape with %2!
    '47 -  %1 flinched!
    '48 -  But it had no effect!
    '49 -  %1 has no moves left!
    '50 -  %1's %2 is disabled!
    '51 -  %1 can't use the same move twice in a row due to the Torment!
    '52 -  %1 can't use %2 after the Taunt!
    '53 -  %1 can't use the sealed %2!
    '54 -  %1 is fast asleep.
    '55 -  %1 is paralyed! It can't move!
    '56 -  %1 is frozen solid!
    '57 -  %1 is immobilied by love!
    '58 -  It hurt itself in its confusion!
    '59 -  The Mirror Move failed!
    '60 -  But it failed to Spit Up a thing!
    '61 -  But it failed to Swallow a thing!
    '62 -  But there was no PP left for the move!
    '63 -  %1's %2 won't go higher!
    '64 -  %1's %2 won't go lower!
    '65 -  %1's stats won't go any higher!
    '66 -  %1's stats won't go any lower!
    '
    '******CHARGERS******
    '67 -  %1 whipped up a whirlwind!
    '68 -  %1 took in sunlight!
    '69 -  %1 lowered its head!
    '70 -  %1 is glowing!
    '71 -  %1 flew up high!
    '72 -  %1 dug a hole!
    '73 -  %1 hid underwater!
    '74 -  %1 sprang up!
    '75 -  %1 is storing energy!
    '76 -  %1 unleashed energy!
    '77 -  %1 must recharge!
    '
    '******BARRIERS AND RECOVERY******
    '78 -  %1's team is covered by a veil!
    '79 -  %1 is protected by %2!
    '80 -  %1 regained health!
    '81 -  %1's HP is full!
    '82 -  %1's team's %2 raised %3!
    '83 -  %1's team's %2 raised %3 a little!
    '84 -  %1 went to sleep!
    '85 -  %1 slept and became healthy!
    '86 -  %1 made a Wish!
    '87 -  %1's Wish came true!
    '88 -  %1 planted its roots!
    '89 -  %1 absorbed nutrients with its roots!
    '90 -  A soothing aroma wafted through the area!
    '91 -  A bell chimed!
    '92 -  All stat changes were eliminated!
    '93 -  %1 braced itself!
    '94 -  %1 Endured the hit!
    '95 -  %1's team became shrouded in Mist!
    '96 -  %1 is protected by Mist!
    '97 -  %1 protected itself!
    '98 -  %1 made a Substitute!
    '99 -  %1 already has a Substitute!
    '100 - The Substitute took damage for %1!
    '101 - %1's Substitute faded!
    '102 - It was too weak to make a Substitute!
    '
    '******WEATHER******
    '103 - It started to rain!
    '104 - A downpour started!
    '105 - Rain continues to fall.
    '106 - The downpour continues.
    '107 - The rain stopped.
    '108 - A sandstorm brewed!
    '109 - The sandstorm rages.
    '110 - The sandstorm subsided.
    '111 - %1 is buffeted by the sandstorm!
    '112 - The sunlight got bright!
    '113 - The sunlight is strong.
    '114 - The sunlight faded.
    '115 - It started to hail!
    '116 - Hail continues to fall.
    '117 - The hail stopped.
    '118 - %1 is stricken by Hail!
    '
    '******ITEMS******
    '119 - %1's %2 cured paralysis!
    '120 - %1's %2 cured poison!
    '121 - %1's %2 healed its burn!
    '122 - %1's %2 defrosted it!
    '123 - %1's %2 woke it from its sleep!
    '124 - %1's %2 snapped it out of confusion!
    '125 - %1's %2 activated!
    '126 - %1's %2 normalied its status!
    '127 - %1's %2 restored health!
    '128 - %1's %2 restored %3's PP!
    '129 - %1's %2 restored its status!
    '130 - %1's %2 restored its HP a little!
    '131 - %1 allows the use of only %2!
    '132 - %1 hung on using its %2!
    '
    '******TRAITS******
    '133 - %1's %2 raised its Speed!
    '134 - %1 was protected by %2!
    '135 - %1's %2 prevents %3 from using %4!
    '136 - %1 restored HP using its %2!
    '137 - %1's %2 made it rain!
    '138 - %1's %2 intensified the sun's rays!
    '139 - %1's %2 whipped up a sandstorm!
    '140 - %1's %2 made %3 useless!
    '141 - %1's %2 made it the %3 type!
    '142 - %1's %2 prevents paralysis!
    '143 - %1's %2 prevents burns!
    '144 - %1's %2 prevents poisoning!
    '145 - %1's %2 prevents confusion!
    '146 - %1's %2 prevents romance!
    '147 - %1's %2 prevents flinching!
    '148 - %1's %2 prevents stat loss!
    '149 - %1's %2 prevents %3 loss!
    '150 - %1's %2 raised its Fire power!
    '151 - %1 anchors itself with %2!
    '152 - %1's %2 cuts %3's Attack!
    '153 - %1's %2 hurt %3!
    '154 - %1 Traced %2's %3!
    '155 - %1's %2 blocks %3!
    '156 - %1's %2 restored some HP!
    '157 - %1's %2 infatuated %3!
    '158 - %1's %2 made %3 ineffective!
    '159 - %1's %2 cured its poison problem!
    '160 - %1 sucked up the %2!
    '161 - %1 transformed!
    '162 - %1's %2 took the attack!
    '163 - %1's %2 prevents switching!
    '164 - %1's %2 prevented %3's %4 from working!
    '165 - %1's %2 made it ineffective!
    '166 - %1's %2 prevents %3's %4 from working!
    '167 - %1's %2 cured its burn problem!
    '168 - %1's %2 had no effect on %3!
    '
    '******MAJOR CONDITIONS******
    '169 - %1 fell asleep!
    '170 - %1's %2 made %3 sleep!
    '171 - %1 is already asleep!
    '172 - %1 wasn't affected!
    '173 - %1 woke up!
    '174 - %1 was poisoned!
    '175 - %1's %2 poisoned %3!
    '176 - %1 is hurt by poison!
    '177 - %1 is already poisoned.
    '178 - %1 was badly poisoned!
    '179 - %1 was burned!
    '180 - %1's %2 burned %3!
    '181 - %1 is hurt by its burn!
    '182 - %1 already has a burn.
    '183 - %1 was frozen solid!
    '184 - %1's %2 froze %3 solid!
    '185 - %1 was defrosted!
    '186 - %1 was defrosted by %2!
    '187 - %1 is paralyed! It may be unable to move!
    '188 - %1's %2 paralyed %3!  It may be unable to move!
    '189 - %1 is already paralyed!
    '190 - %1 was healed of paralysis!
    '
    '******MINOR CONDITIONS******
    '191 - %1 is confused!
    '192 - %1 snapped out of confusion!
    '193 - %1 became confused!
    '194 - %1 is already confused!
    '195 - %1 became confused out of fatigue!
    '196 - %1 fell in love!
    '197 - %1 is in love with %2!
    '198 - %1 fell into a Nightmare!
    '199 - %1 is locked in a Nightmare!
    '200 - %1 cut its own HP and laid a Curse on %2!
    '201 - %1 is afflicted by the Curse!
    '202 - %1 was squeeed by %2's %3!
    '203 - %1 was trapped in the vortex!
    '204 - %1 was trapped by %2!
    '205 - %1 was Wrapped by %2!
    '206 - %1 Clamped %2!
    '207 - %1 is hurt by %2!
    '208 - %1 was freed from %2!
    '209 - %1 was seeded!
    '210 - %1 evaded the attack!
    '211 - %1's health is sapped by %2!
    '
    '******MISC******
    '212 - %1 is getting pumped!
    '213 - Magnitude %1!
    '214 - The Present healed %1!
    '215 - %1 had its energy drained!
    '216 - %1's dream was eaten!
    '217 - %1's team's %2 stopped working!
    '218 - %1's team's %2 faded!
    '219 - %1 was blown away!
    '220 - %1 transformed into the %2 type!
    '221 - %1 kept going and crashed!
    '222 - But %1's Uproar kept it awake!
    '223 - %1 woke up in the Uproar!
    '224 - %1 caused an Uproar!
    '225 - %1 is making an Uproar!
    '226 - %1 calmed down.
    '227 - But %1 can't sleep in an Uproar!
    '228 - %1 learned %2!
    '229 - But the Uproar kept %1 awake!
    '230 - %1 stayed awake using its %2!
    '231 - %1 is unaffected!
    '232 - %1 transformed into %2!
    '233 - %1's Rage is building!
    '234 - %1's %2 was disabled!
    '235 - %1 is disabled no more!
    '236 - %1 got an Encore!
    '237 - %1's Encore ended!
    '238 - %1 took aim at %2!
    '239 - %1 Sketched %2!
    '240 - %1 is trying to take its foe with it!
    '241 - %1 took %2 with it!
    '242 - Reduced %1's %2's PP by %3!
    '243 - %1 stole %2's %3!
    '244 - %1 can't escape now!
    '245 - %1's team had Spikes scattered around it!
    '246 - %1 is hurt by Spikes!
    '247 - %1 identified %2!
    '248 - %1's Perish count fell to %2!
    '249 - %1 cut its own HP and maximied Attack!
    '250 - %1 copied %2's stat changes!
    '251 - %1 got free of %2's %3!
    '252 - %1 shed Leech Seed!
    '253 - %1 blew away Spikes!
    '254 - %1 fled from battle!
    '255 - %1 foresaw an attack!
    '256 - %1 took %2's attack!
    '257 - %1 chose %2 as its destiny!
    '258 - %1 became the center of attention!
    '259 - %1 began charging power!
    '260 - %1 turned into %2!
    '261 - %1's status returned to normal!
    '262 - %1 was subjected to Torment!
    '263 - %1 is tightening its focus!
    '264 - %1 fell for the Taunt!
    '265 - %1 is ready to help %2!
    '266 - %1 switched items with %2!
    '267 - %1 obtained %2.
    '268 - %1 copied %2's %3!
    '269 - %1 anchored itself with its roots!
    '270 - %1 made %2 drowsy!
    '271 - %1 knocked off %2's %3!
    '272 - %1 swapped abilities with %2!
    '273 - %1 sealed the opponent's moves!
    '274 - %1 wants the opponent to bear a Grudge!
    '275 - %1's %2 lost all its PP due to the Grudge!
    '276 - %1 shrouded itself in %2!
    '277 - %1's %2 was bounced back by %3!
    '278 - %1 awaits its foe to make a move!
    '279 - %1 Snatched %2's move!
    '280 - Electricity's power was weakened!
    '281 - Fire's power was weakened!
    '282 - %1 found one %2!
    '283 - %1 used %2 to hustle!
    '284 - %1 lost its focus and couldn't move!
    '285 - %1 (Lv.%2 %3) was dragged out!
    '286 - The wall shattered!
    '287 - %1 Stockpiled %2!
    '288 - %1 can't Stockpile any more!
    '289 - The wind turned into a Heat Wave!
    '290 - Coins scattered everywhere!
    '291 - The battlers shared their pain!
    '292 - All affected Pokémon will faint in 3 turns!

    '******LATE ADDITIONS******
    '293 - %1 is loafing around!
    '294 - %1's attack!
    '295 - %1's attack continues!
    '296 - %1 became shrouded in Mist!
    '297 - %1 withdrew %2!
    '298 - %1 sent out %2 (Lv.%3 %4)!
    '299 - %1's %2 cured its sleep problem!
    '300 - %1's %2 cured its paralysis problem!
    '301 - %1's %2 cured its ice problem!
    '302 - %1's %2 cured its confusion problem!
    '303 - %1's %2 cured its love problem!
    '304 - Using %1, the %2 of %3 rose!
    '305 - Using %1, the %2 of %3 sharply rose!
    '306 - For %1, %2 was too spicy!
    '307 - For %1, %2 was too sour!
    '308 - For %1, %2 was too sweet!
    '309 - For %1, %2 was too dry!
    '310 - For %1, %2 was too bitter!
    '311 - There's no PP left for this move!
    '312 - Battle Mode: Red/Blue/Yellow
    '313 - Battle Mode: Gold/Silver/Crystal
    '314 - Battle Mode: Ruby/Sapphire
    '315 - ---
    
    ModeText(0) = "RBY (Trades)"
    ModeText(1) = "GSC (Trades)"
    ModeText(2) = "Ru/Sa Only"
    ModeText(3) = "Full Advance"
    ModeText(4) = "Advance (Trades)"
    ModeText(5) = "True RBY"
    ModeText(6) = "True GSC"
    
    'These are for Pokedex and similar displays.
    ModeText(7) = "Red"
    ModeText(8) = "Blue"
    ModeText(9) = "Yellow"
    ModeText(10) = "Gold"
    ModeText(11) = "Silver"
    ModeText(12) = "Crystal"
    ModeText(13) = "Ruby"
    ModeText(14) = "Sapphire"
    
    TerrainText(nbStadium) = "Stadium"
    TerrainText(nbShortGrass) = "Short Grass"
    TerrainText(nbTallGrass) = "Tall Grass"
    TerrainText(nbVeryTallGrass) = "Very Tall Grass"
    TerrainText(nbOcean) = "Ocean"
    TerrainText(nbUnderwater) = "Underwater"
    TerrainText(nbPond) = "Pond"
    TerrainText(nbsAnd) = "Sand"
    TerrainText(nbCave) = "Cave"
    TerrainText(nbMountain) = "Rock"
    
    ProgramText(1) = "Pokémon NetBattle"
    ProgramText(2) = "&File"
    ProgramText(3) = "E&xit"
    ProgramText(4) = "&Save"
    ProgramText(5) = "&Load"
    ProgramText(6) = "&Boxes >>"
    ProgramText(7) = "<< &Boxes"
    ProgramText(8) = "&Done"
    ProgramText(9) = "&Arrange"
    ProgramText(10) = "Image"
    ProgramText(11) = "Info"
    ProgramText(12) = "Graphics"
    ProgramText(13) = "User Name"
    ProgramText(14) = "Auto Messages"
    ProgramText(15) = "Win"
    ProgramText(16) = "Lose"
    ProgramText(17) = "Trainer"
    ProgramText(18) = "PKMN %1"
    ProgramText(19) = "Compat."
    ProgramText(20) = "Team Builder"
    ProgramText(21) = "Box %1"
    ProgramText(22) = "Pokémon Boxes"
    ProgramText(23) = "&Sort"
    ProgramText(24) = "By &PokéDex Number"
    ProgramText(25) = "By &GSC PokéDex"
    ProgramText(26) = "By &Advance PokéDex"
    ProgramText(27) = "By &Name"
    ProgramText(28) = "&Window"
    ProgramText(29) = "&Help"
    ProgramText(30) = "&Web Site"
    ProgramText(31) = "&About"
    ProgramText(32) = "&New"
    ProgramText(33) = "&Open"
    ProgramText(34) = "Save &As"
    ProgramText(35) = "&Print"
    ProgramText(36) = "&Close Team Builder"
    ProgramText(37) = "Your Team"
    ProgramText(38) = "Pokémon"
    ProgramText(39) = "Nickname"
    ProgramText(40) = "Item"
    ProgramText(41) = "Moves"
    ProgramText(42) = "Type1"
    ProgramText(43) = "Type2"
    ProgramText(44) = "Current Moves:"
    ProgramText(45) = "Overwrite the current Pokémon?"
    ProgramText(46) = "Confirm Change"
    ProgramText(47) = "You already have a %1 on your team!"
    ProgramText(48) = "Duplicate Pokémon"
    ProgramText(49) = "Are you sure you want to delete this Pokémon?"
    ProgramText(50) = "Confirm Delete"
    ProgramText(51) = "%1  Exit anyway?"
    ProgramText(52) = "Not Ready"
    ProgramText(53) = "Save Trainer/Team"
    ProgramText(54) = "Pokémon NetBattle File %1"
    ProgramText(55) = "Names cannot contain commas, colons, or apostrophes!%nPlease edit the appropriate names and try again."
    ProgramText(56) = "Error"
    ProgramText(57) = "Do you want to save the changed team?%nPicking No will revert your team to the way it was when you opened the Team Builder.%nPicking Cancel will close the Team Builder, but keep new changes."
    ProgramText(58) = "Save Changes"
    ProgramText(59) = "Ready to Battle!"
    ProgramText(60) = "Your current team is not %1 compatible.  Some Pokémon/moves will be reset.  Would you like to continue?"
    ProgramText(61) = "Version Change"
    ProgramText(62) = "%1 was added in %2"
    ProgramText(63) = "This is learned by %1 %2"
    ProgramText(64) = "Level"
    ProgramText(65) = "Breeding"
    ProgramText(66) = "Special"
    ProgramText(67) = "Move Tutor"
    ProgramText(68) = "Acc: %1% - %2"
    ProgramText(69) = "Level"
    ProgramText(70) = "Held Item"
    ProgramText(71) = "Max HP"
    ProgramText(72) = "Move"
    ProgramText(73) = "Pokémon can only use four moves at a time!"
    ProgramText(74) = "That move doesn't work right yet - you can pick it, but it's special ability might not work."
    ProgramText(75) = "Warning"
    ProgramText(76) = "Illegal Move"
    '>>> Call WriteDebugLog("Set complete.")
End Sub

Public Sub OpenDebugLog()
    Dim X As Integer
    If Not DoDebugLogs Then Exit Sub
    DebugLogName = "debug " & IIf(Command$ = "SERVER", "server", "client") & ".log"
    X = FreeFile
    Open SlashPath & DebugLogName For Output As #X
    Print #X, "Log started " & Date & " at " & Time
    Print #X, "NetBattle v" & App.Major & "." & App.Minor & "." & BetaRel
    If InVBMode Then Print #X, "(Running inside Visual Basic)"
    Close #X
End Sub

Public Sub CloseDebugLog()
    Dim X As Integer
    If Not DoDebugLogs Then Exit Sub
    X = FreeFile
    Open SlashPath & DebugLogName For Append As #X
    Print #X, Time, "Log ended normally."
    Close #X
End Sub

Public Sub WriteDebugLog(Text As String)
    Dim X As Integer
    On Error GoTo Failed
    If Not DoDebugLogs Then Exit Sub
    X = FreeFile
    Open SlashPath & DebugLogName For Append As #X
    Print #X, Time, Timer, Text
    Close #X
Failed:
End Sub
Public Function CompatVersion(ByVal Compat As CompatModes) As BattleModes
    Select Case Compat
    Case 0, 5: CompatVersion = nbRBYBattle
    Case 1, 6: CompatVersion = nbGSCBattle
    Case 2, 3, 4, 7: CompatVersion = nbAdvBattle
    End Select
End Function
Public Function AdvItem(ItemNum As Items) As Boolean
    Select Case ItemNum
    Case 0, 13 To 29, 31, 33 To UBound(Item)
        AdvItem = True
    Case Else
        AdvItem = False
    End Select
End Function

Public Function AdvItem2(ItemNum As Integer) As Boolean
    Select Case ItemNum
    Case 0, 13 To 29, 31, 33 To UBound(Item)
        AdvItem2 = True
    Case Else
        AdvItem2 = False
    End Select
End Function

'It looks big and complicated for such a simple task, but this
'string sorting algorythm is the fastest I've seen.  Taken from
'Philippe Lord's mdlArray.bas, avaliable at Planet Source Code.

Public Sub SortStringArray(ByRef SArray() As String)
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(SArray)
   iUBound = UBound(SArray)
   TriQuickSortString2 SArray, 4, iLBound, iUBound
   InsertionSortString SArray, iLBound, iUBound
End Sub

Private Sub TriQuickSortString2(ByRef SArray() As String, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)
   Dim i     As Long
   Dim J     As Long
   Dim sTemp As String
   On Error GoTo ETrap
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2
      If SArray(iMin) > SArray(i) Then SwapStrings SArray(iMin), SArray(i)
      If SArray(iMin) > SArray(iMax) Then SwapStrings SArray(iMin), SArray(iMax)
      If SArray(i) > SArray(iMax) Then SwapStrings SArray(i), SArray(iMax)
      J = iMax - 1
      SwapStrings SArray(i), SArray(J)
      i = iMin
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(SArray(J)), 4
      Do
         Do
            i = i + 1
         Loop While SArray(i) < sTemp
         Do
            J = J - 1
         Loop While SArray(J) > sTemp
         If J < i Then Exit Do
         SwapStrings SArray(i), SArray(J)
      Loop
      SwapStrings SArray(i), SArray(iMax - 1)
      TriQuickSortString2 SArray, iSplit, iMin, J
      TriQuickSortString2 SArray, iSplit, i + 1, iMax
   End If
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
   Exit Sub
ETrap:
    DoEvents
    Resume
End Sub

Private Sub InsertionSortString(ByRef SArray() As String, ByVal iMin As Long, ByVal iMax As Long)
   Dim i As Long
   Dim J As Long
   Dim sTemp As String
      For i = iMin + 1 To iMax
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(SArray(i)), 4
      J = i
      Do While J > iMin
         If SArray(J - 1) <= sTemp Then Exit Do
         CopyMemory ByVal VarPtr(SArray(J)), ByVal VarPtr(SArray(J - 1)), 4
         J = J - 1
      Loop
      CopyMemory ByVal VarPtr(SArray(J)), ByVal VarPtr(sTemp), 4
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

'Number processing
'These have been replaced with new uber-speedy code
Public Function Dec2Bin(ByVal X As Long, ByVal Fixed As Integer) As String
    Static lDone As Long
    Static sByte(0 To 255) As String
    Dim sNibble(0 To 15) As String
    Dim Y As Long
    If Sgn(X) = -1 And InVBMode Then Stop
    If lDone = 0 Then
        sNibble(0) = "0000"
        sNibble(1) = "0001"
        sNibble(2) = "0010"
        sNibble(3) = "0011"
        sNibble(4) = "0100"
        sNibble(5) = "0101"
        sNibble(6) = "0110"
        sNibble(7) = "0111"
        sNibble(8) = "1000"
        sNibble(9) = "1001"
        sNibble(10) = "1010"
        sNibble(11) = "1011"
        sNibble(12) = "1100"
        sNibble(13) = "1101"
        sNibble(14) = "1110"
        sNibble(15) = "1111"
        For lDone = 0 To 255
            sByte(lDone) = sNibble(lDone \ &H10) & sNibble(lDone And &HF)
        Next
    End If
    
    If X < &H100 Then
        Dec2Bin = Right$(sByte(X), Fixed)
    ElseIf X < &H10000 Then
        Dec2Bin = Right$( _
                  sByte(X \ &H100 And &HFF) & _
                  sByte(X And &HFF), Fixed)
    ElseIf X < &H1000000 Then
        Dec2Bin = Right$( _
                  sByte(X \ &H10000 And &HFF) & _
                  sByte(X \ &H100 And &HFF) & _
                  sByte(X And &HFF), Fixed)
    Else
        Dec2Bin = Right$( _
                  sByte(X \ &H1000000 And &HFF) & _
                  sByte(X \ &H10000 And &HFF) & _
                  sByte(X \ &H100 And &HFF) & _
                  sByte(X And &HFF), Fixed)
    End If
    Y = Len(Dec2Bin)
    If Y < Fixed Then Dec2Bin = String(Fixed - Y, "0") & Dec2Bin
    
    If InVBMode Then
        If X <> Bin2Dec(Dec2Bin) Then Err.Raise 6
    End If

End Function

Public Function Bin2Dec(ByVal BitString As String) As Long
    Dim X As Long
    Static T() As Integer
    If BitString = vbNullString Then Exit Function
    ReDim T(0 To Len(BitString) - 1)
    CopyMemory T(0), ByVal StrPtr(BitString), LenB(BitString)
    Bin2Dec = T(0) - vbKey0
    For X = 1 To UBound(T)
        Bin2Dec = Bin2Dec + Bin2Dec + T(X) - vbKey0
    Next X
End Function

Public Function Bin2Chr(ByVal BinString As String) As String
    Dim X As Long
    Dim Y As Long
    Dim Build As String
    'Make the string length a multiple of 8
    X = 8 - Len(BinString) Mod 8
    If X = 8 Then X = 0
    BinString = BinString & String$(X, "0")
    Y = Len(BinString) \ 8
    'Fill the buffer
    Build = String$(Y, vbNullChar)
    'Take each 8 digit set of binary and convert to decimal, and then
    'convert it to a character.
    For X = 1 To Y
        Mid(Build, X) = Chr$(Bin2Dec(Mid(BinString, X * 8 - 7, 8)))
    Next X
    Bin2Chr = Build
End Function

Public Function Chr2Bin(ByVal ChrString As String) As String
    Dim Build As String
    Dim X As Long
    'Reverse of the above
    Build = String$(Len(ChrString) * 8, vbNullChar)
    For X = 1 To Len(ChrString)
        Mid(Build, X * 8 - 7) = Dec2Bin(Asc(Mid(ChrString, X, 1)), 8)
    Next X
    Chr2Bin = Build
End Function
Public Function Int2Str(ByVal Number As Integer) As String
    Dim Temp As String
    Temp = "  "
    CopyMemory ByVal Temp, Number, ByVal 2
    Int2Str = Temp
End Function
Public Function Str2Int(ByVal iString As String) As Integer
    Dim X As Integer
    If Len(iString) <> 2 Then Err.Raise 6
    CopyMemory X, ByVal iString, 2
    Str2Int = X
End Function
Public Function Lng2Str(ByVal Number As Long) As String
    Dim Temp As String
    Temp = "    "
    CopyMemory ByVal Temp, Number, ByVal 4
    Lng2Str = Temp
End Function
Public Function Str2Lng(ByVal iString As String) As Long
    Dim X As Long
    If Len(iString) <> 4 Then Err.Raise 6
    CopyMemory X, ByVal iString, 4
    Str2Lng = X
End Function
Public Function Hex2Chr(ByVal HexString As String) As String
    Dim X As Long
    Dim Y As Long
    Hex2Chr = String$(Len(HexString) \ 2, vbNullChar)
    For X = 1 To Len(HexString) Step 2
        Y = Y + 1
        Y = Y + 1
        Mid(Hex2Chr, Y, 1) = Chr$(Val("&H" & Mid$(HexString, X, 2)))
    Next X
End Function

Function GetFText(ByVal StringNum As Integer) As String
    Dim X As Integer
    Dim Temp As String
    X = FreeFile
    Temp = String$(FTextLen(StringNum), vbNullChar)
    Open FTextFile For Binary Access Read As #X
    Get #X, FTextOffset(StringNum), Temp
    Close #X
    If StringNum = 202 Then 'Note to self: fix this later
        Temp = Replace(Temp, "squeeed", "squeezed")
    End If
    
    GetFText = Temp
End Function

Function PText(ByVal StringNum As Integer, _
    Optional ByVal Param1 As String = "", _
    Optional ByVal Param2 As String = "", _
    Optional ByVal Param3 As String = "", _
    Optional ByVal Param4 As String = "") As String
    
    Dim Temp As String
    Dim X As Integer
    
    Temp = ProgramText(StringNum)
    Temp = Replace(Temp, "%n", vbCrLf)
    If Not Param1 = "" Then
        Temp = Replace(Temp, "%1", Param1)
    End If
    If Not Param2 = "" Then
        Temp = Replace(Temp, "%2", Param2)
    End If
    If Not Param3 = "" Then
        Temp = Replace(Temp, "%3", Param3)
    End If
    If Not Param4 = "" Then
        Temp = Replace(Temp, "%4", Param4)
    End If
    PText = Temp
End Function

Public Function SleepIt(ByVal Milliseconds As Long)
    Dim T1 As Single
    Dim T2 As Single
    Dim T3 As Single
    T3 = Timer
    T1 = T3 + (Milliseconds / 1000)
    Do
        T2 = Timer
        If T2 < T3 Then T2 = T2 + 86400 'In case it happens to be midnight
        Sleep 1 'Free up processing
        DoEvents
    Loop Until T2 >= T1
End Function
Public Function PKMN2Str(Poke As Pokemon) As String '35 bytes
    Dim Build As String
    Dim X As Integer
    With Poke
        X = .Gender - 1
        If X = -1 Then X = 0
        Build = Dec2Bin(.No, 9) & _
        Dec2Bin(.GameVersion, 3) & _
        Dec2Bin(.Level, 7) & _
        Dec2Bin(.Item, 7) & _
        Dec2Bin(.NatureNum, 5) & _
        CStr(.AttNum) & _
        CStr(X) & _
        Bool2Bin(.Shiny) & _
        Dec2Bin(.InBox, 4) & _
        Dec2Bin(.UnownLetter, 5) & _
        Dec2Bin(.Move(1), 9) & _
        Dec2Bin(.Move(2), 9) & _
        Dec2Bin(.Move(3), 9) & _
        Dec2Bin(.Move(4), 9) & _
        Dec2Bin(.DV_HP, 5) & Dec2Bin(.DV_Atk, 5) & _
        Dec2Bin(.DV_Def, 5) & Dec2Bin(.DV_Spd, 5) & _
        Dec2Bin(.DV_SAtk, 5) & Dec2Bin(.DV_SDef, 5) & _
        Dec2Bin(.EV_HP, 8) & Dec2Bin(.EV_Atk, 8) & _
        Dec2Bin(.EV_Def, 8) & Dec2Bin(.EV_Spd, 8) & _
        Dec2Bin(.EV_SAtk, 8) & Dec2Bin(.EV_SDef, 8)
        PKMN2Str = Pad(.Nickname, 15) & Bin2Chr(Build)
        'Debug.Print Dec2Bin(.Move(1), 9)
        'Debug.Print Dec2Bin(.Move(2), 9)
        'Debug.Print Dec2Bin(.Move(3), 9)
        'Debug.Print Dec2Bin(.Move(4), 9)
        'Debug.Print "Encoded PKMN:", .No, .Move(1), .Move(2), .Move(3), .Move(4)
        'Debug.Print Len(Build), Len(Bin2Chr(Build))
    End With
End Function
Public Function Str2PKMN(PokeStr As String, Optional SkipFillIn As Boolean = False) As Pokemon
    Dim Build As Pokemon
    Dim Temp As String
    Dim X As Integer
    Temp = Trim(ChopString(PokeStr, 15))
    PokeStr = Chr2Bin(PokeStr)
    X = Bin2Dec(ChopString(PokeStr, 9))
    If X > UBound(BasePKMN) Or X = 0 Then Exit Function
    Build = BasePKMN(X)
    With Build
        .Nickname = Temp
        .GameVersion = Bin2Dec(ChopString(PokeStr, 3))
        .Level = Bin2Dec(ChopString(PokeStr, 7))
        .Item = Bin2Dec(ChopString(PokeStr, 7))
        .NatureNum = Bin2Dec(ChopString(PokeStr, 5))
        .AttNum = Val(ChopString(PokeStr, 1))
        .Gender = Bin2Dec(ChopString(PokeStr, 1)) + 1
        .Shiny = CBool(ChopString(PokeStr, 1))
        .InBox = Bin2Dec(ChopString(PokeStr, 4))
        .UnownLetter = Bin2Dec(ChopString(PokeStr, 5))
        .Move(1) = Bin2Dec(ChopString(PokeStr, 9))
        .Move(2) = Bin2Dec(ChopString(PokeStr, 9))
        .Move(3) = Bin2Dec(ChopString(PokeStr, 9))
        .Move(4) = Bin2Dec(ChopString(PokeStr, 9))
        .DV_HP = Bin2Dec(ChopString(PokeStr, 5))
        .DV_Atk = Bin2Dec(ChopString(PokeStr, 5))
        .DV_Def = Bin2Dec(ChopString(PokeStr, 5))
        .DV_Spd = Bin2Dec(ChopString(PokeStr, 5))
        .DV_SAtk = Bin2Dec(ChopString(PokeStr, 5))
        .DV_SDef = Bin2Dec(ChopString(PokeStr, 5))
        .EV_HP = Bin2Dec(ChopString(PokeStr, 8))
        .EV_Atk = Bin2Dec(ChopString(PokeStr, 8))
        .EV_Def = Bin2Dec(ChopString(PokeStr, 8))
        .EV_Spd = Bin2Dec(ChopString(PokeStr, 8))
        .EV_SAtk = Bin2Dec(ChopString(PokeStr, 8))
        .EV_SDef = Bin2Dec(ChopString(PokeStr, 8))
    End With
    If Not SkipFillIn Then Call FillInPokeData(Build, Build.GameVersion)
    Str2PKMN = Build
End Function
Public Sub FillInPokeData(Poke As Pokemon, ByVal GameVersion As Byte)
    Dim X As Integer
    With Poke
        .GameVersion = GameVersion
        If .Nickname = "" Then .Nickname = .Name
        If .Level > 100 Or .Level = 0 Then .Level = 100
        If .Item > UBound(Item) Then
            .Item = nbNoItem
        ElseIf Item(.Item) = "" Then
            .Item = nbNoItem
        End If
        
        '.Image = ChooseImage(.No, .DV_Atk, .DV_Def, .DV_Spd, .DV_SAtk, 1)

        Select Case GameVersion
        Case 2, 3, 4
            If Not AdvItem(.Item) Then .Item = nbNoItem
            X = .EV_HP
            X = X + .EV_Atk
            X = X + .EV_Def
            X = X + .EV_Spd
            X = X + .EV_SAtk
            X = X + .EV_SDef
            If X > 510 Then
                .EV_HP = 85: .EV_Atk = 85: .EV_Def = 85: .EV_Spd = 85: .EV_SAtk = 85: .EV_SDef = 85
            End If
            .ModAttr(0) = BasePKMN(.No).ModAttr(0)
            .ModAttr(1) = BasePKMN(.No).ModAttr(1)
            
            If .GameVersion = nbModAdv Then
                If .AttNum = 1 And .ModAttr(1) = nbNoTrait Then .AttNum = 0
                .Attribute = .ModAttr(.AttNum)
            Else
                If .AttNum = 1 And .PAtt(1) = nbNoTrait Then .AttNum = 0
                .Attribute = .PAtt(.AttNum)
            End If
            If .NatureNum > 24 Then .NatureNum = 0
            .MaxHP = GetAdvHP(.BaseHP, .DV_HP, .EV_HP, .Level)
            .HP = .MaxHP
            .Attack = GetAdvStat(.BaseAttack, .DV_Atk, .EV_Atk, .Level, Nature(.NatureNum).StatChg(1))
            .Defense = GetAdvStat(.BaseDefense, .DV_Def, .EV_Def, .Level, Nature(.NatureNum).StatChg(2))
            .Speed = GetAdvStat(.BaseSpeed, .DV_Spd, .EV_Spd, .Level, Nature(.NatureNum).StatChg(3))
            .SpecialAttack = GetAdvStat(.BaseSAttack, .DV_SAtk, .EV_SAtk, .Level, Nature(.NatureNum).StatChg(4))
            .SpecialDefense = GetAdvStat(.BaseSDefense, .DV_SDef, .EV_SDef, .Level, Nature(.NatureNum).StatChg(5))
            Select Case .PercentFemale
                Case -1: .Gender = 0
                Case 0: .Gender = 1
                Case 16: .Gender = 2
            End Select
            If .UnownLetter > 27 Then .UnownLetter = 0
        Case Else
            If .DV_Atk > 15 Then .DV_Atk = 15
            If .DV_Def > 15 Then .DV_Def = 15
            If .DV_Spd > 15 Then .DV_Spd = 15
            If .DV_SAtk > 15 Then .DV_SAtk = 15
            .DV_SDef = 0
            .Attack = GetStat(.Level, .BaseAttack, .DV_Atk)
            .Defense = GetStat(.Level, .BaseDefense, .DV_Def)
            .Speed = GetStat(.Level, .BaseSpeed, .DV_Spd)
            Select Case GameVersion
            Case 0, 5
                .SpecialAttack = GetStat(.Level, .BaseSpecial, .DV_SAtk)
                .SpecialDefense = GetStat(.Level, .BaseSpecial, .DV_SAtk)
                .Item = nbNoItem
                .Shiny = False
            Case Else
                .SpecialAttack = GetStat(.Level, .BaseSAttack, .DV_SAtk)
                .SpecialDefense = GetStat(.Level, .BaseSDefense, .DV_SAtk)
                .Shiny = (ShinyDV(.DV_Atk) And .DV_Def = 10 And .DV_Spd = 10 And .DV_SAtk = 10)
            End Select
            .DV_HP = 0
            If .DV_Atk Mod 2 = 1 Then .DV_HP = .DV_HP + 8
            If .DV_Def Mod 2 = 1 Then .DV_HP = .DV_HP + 4
            If .DV_Spd Mod 2 = 1 Then .DV_HP = .DV_HP + 2
            If .DV_SAtk Mod 2 = 1 Then .DV_HP = .DV_HP + 1
            .MaxHP = GetHP(.Level, .BaseHP, .DV_HP)
            .HP = .MaxHP
            If .PercentFemale = -1 Then
                .Gender = 0
            Else
                If .DV_Atk <= .PercentFemale - 1 Then .Gender = 2 Else .Gender = 1
            End If
            .Attribute = nbNoTrait
            If .Item > 41 Then .Item = nbNoItem
            If .No = 201 Then
                .UnownLetter = (( _
                ((.DV_Atk Mod 8) \ 2) * 64 + _
                ((.DV_Def Mod 8) \ 2) * 16 + _
                ((.DV_Spd Mod 8) \ 2) * 4 + _
                (.DV_SAtk Mod 8) \ 2) \ 10) Mod 26
            Else
                .UnownLetter = 0
            End If
        End Select
    End With
End Sub

Public Sub ResetDefaultSound(ByVal SoundNumber As Byte)
    Select Case SoundNumber
        Case nbSoundOpening
            SoundFile(nbSoundOpening) = SlashPath & "Media\opening.wav"
        Case nbSoundSignon
            SoundFile(nbSoundSignon) = SlashPath & "Media\signon.wav"
        Case nbSoundChat
            SoundFile(nbSoundChat) = SlashPath & "Media\chat.wav"
        Case nbSoundChallenge
            SoundFile(nbSoundChallenge) = SlashPath & "Media\challenge.wav"
        Case nbMusicOpening
            SoundFile(nbMusicOpening) = SlashPath & "Media\opening.mid"
        Case nbMusicGSC
            SoundFile(nbMusicGSC) = SlashPath & "Media\GSBattle.it"
        Case nbMusicRBY
            SoundFile(nbMusicRBY) = SlashPath & "Media\RBYBattle.it"
        Case nbMusicVictory
            SoundFile(nbMusicVictory) = SlashPath & "Media\victory.it"
        Case nbMusicLost
            SoundFile(nbMusicLost) = SlashPath & "Media\lost.mid"
        Case nbMusicChallenge
            SoundFile(nbMusicChallenge) = SlashPath & "Media\challenge.it"
        Case nbMusicRuSa
            SoundFile(nbMusicRuSa) = SlashPath & "Media\RSBattle.it"
    End Select
End Sub

Public Sub MakeMoveArray(ByVal PokeNo As Integer, ByVal GameVersion As CompatModes, TempMove() As Integer, TempSource() As String)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim B As Boolean
    Dim Temp As String
    Dim TempPKMN As Pokemon
    Dim AdvMoves As Long
    
    'Okay, adding comments so I can try and figure out what the heck is going on here.
    With BasePKMN(PokeNo)
        ReDim TempMove(0)
        'Start with Advance modes, because they're easier.
        If CompatVersion(GameVersion) = nbAdvBattle Then
            If GameVersion = nbModAdv Then AdvMoves = UBound(.AdvMoves) Else AdvMoves = .TotalAdvMoves
            If GameVersion = nbTrueRuSa Then
                ReDim TempMove(1 To AdvMoves + UBound(.ADVTM) + UBound(.AdvBreeding) + UBound(.AdvSpecial))
            Else
                ReDim TempMove(1 To AdvMoves + UBound(.ADVTM) + UBound(.AdvBreeding) + UBound(.AdvSpecial) + UBound(.AdvTutor) + UBound(.LFOnly))
            End If
            ReDim TempSource(1 To UBound(TempMove))
            'These are the common Advance moves
            Y = 0
            For X = 1 To AdvMoves
                Y = Y + 1
                TempMove(Y) = .AdvMoves(X)
                If X <= .TotalAdvMoves Then
                    TempSource(Y) = "Level"
                Else
                    TempSource(Y) = "DB Mod"
                End If
            Next X
            For X = 1 To UBound(.ADVTM)
                Y = Y + 1
                TempMove(Y) = .ADVTM(X)
                TempSource(Y) = Moves(.ADVTM(X)).ADVTM
            Next X
            For X = 1 To UBound(.AdvBreeding)
                Y = Y + 1
                TempMove(Y) = .AdvBreeding(X)
                TempSource(Y) = "Egg Move"
            Next X
            For X = 1 To UBound(.AdvSpecial)
                Y = Y + 1
                TempMove(Y) = .AdvSpecial(X)
                'Note: Change this when NYPC Advance or Colosseum Special moves are known
                TempSource(Y) = "Box/NYPC"
            Next X
            'This bit is for Leaf/Fire Stuff
            If GameVersion = nbFullAdvance Or GameVersion = nbModAdv Then
                For X = 1 To UBound(.AdvTutor)
                    Y = Y + 1
                    TempMove(Y) = .AdvTutor(X)
                    TempSource(Y) = "Move Tutor"
                Next X
                For X = 1 To UBound(.LFOnly)
                    Y = Y + 1
                    TempMove(Y) = .LFOnly(X)
                    TempSource(Y) = "Fire/Leaf"
                Next X
            End If
        Else
            'Okay, now let's move on to GSC/RBY
            Select Case GameVersion
                Case nbRBYTrade, nbGSCTrade
                    ReDim TempMove(1 To UBound(.BaseMoves) + UBound(.MachineMoves) + UBound(.BreedingMoves) + UBound(.SpecialMoves) + UBound(.MoveTutor) + UBound(.RBYMoves) + UBound(.RBYTM))
                Case nbTrueRBY
                    ReDim TempMove(1 To UBound(.RBYMoves) + UBound(.RBYTM))
                Case nbTrueGSC
                    ReDim TempMove(1 To UBound(.BaseMoves) + UBound(.MachineMoves) + UBound(.BreedingMoves) + UBound(.MoveTutor))
            End Select
            ReDim TempSource(1 To UBound(TempMove))
            'Start filling in the move array
            'RBY: This is first for RBY-based modes
            If GameVersion = nbTrueRBY Or GameVersion = nbRBYTrade Then
                For X = 1 To UBound(.RBYMoves)
                    Y = Y + 1
                    TempMove(Y) = .RBYMoves(X)
                    TempSource(Y) = "Level"
                Next X
                For X = 1 To UBound(.RBYTM)
                    Y = Y + 1
                    TempMove(Y) = .RBYTM(X)
                    'Surfing Pikachu - Change HM to Stadium
                    If (.No = 25 Or .No = 26) And TempMove(Y) = 216 Then
                        TempSource(Y) = "Stadium"
                    Else
                        TempSource(Y) = Moves(.RBYTM(X)).OldTM
                    End If
                Next X
            End If
            'Amnesia Psyduck - Add in if in True RBY
            If GameVersion = nbTrueRBY Then
                If .No = 54 Or .No = 55 Then
                    ReDim Preserve TempMove(1 To UBound(TempMove) + 1)
                    ReDim Preserve TempSource(1 To UBound(TempSource) + 1)
                    Y = Y + 1
                    TempMove(Y) = 6
                    TempSource(Y) = "Stadium"
                End If
            End If
            'GSC moves for everything EXCEPT True RBY
            If GameVersion <> nbTrueRBY Then
                Temp = IIf(GameVersion = nbRBYTrade, "GSC Level", "Level")
                For X = 1 To UBound(.BaseMoves)
                    Y = Y + 1
                    TempMove(Y) = .BaseMoves(X)
                    TempSource(Y) = Temp
                    'Change Smeargle's non-Sketch moves to "Sketch" (translation-friendly, even)
                    If .No = 235 And TempMove(Y) <> 184 Then TempSource(Y) = Moves(184).Name
                Next X
                Temp = IIf(GameVersion = nbRBYTrade, "GSC ", "")
                For X = 1 To UBound(.MachineMoves)
                    Y = Y + 1
                    TempMove(Y) = .MachineMoves(X)
                    TempSource(Y) = Temp & Moves(.MachineMoves(X)).NewTM
                Next X
                Temp = IIf(GameVersion = nbRBYTrade, "GSC Egg Move", "Egg Move")
                For X = 1 To UBound(.BreedingMoves)
                    If GameVersion = nbTrueGSC Then
                        TempPKMN = BasePKMN(PokeNo)
                        TempPKMN.GameVersion = nbTrueGSC
                        TempPKMN.Move(1) = .BreedingMoves(X)
                        If LegalMove(TempPKMN, True) = "" Then
                            Y = Y + 1
                            TempMove(Y) = .BreedingMoves(X)
                            TempSource(Y) = Temp
                        End If
                    Else
                        Y = Y + 1
                        TempMove(Y) = .BreedingMoves(X)
                        TempSource(Y) = Temp
                    End If
                Next X
                Temp = IIf(GameVersion = nbRBYTrade, "GSC Tutor", "Move Tutor")
                For X = 1 To UBound(.MoveTutor)
                    Y = Y + 1
                    TempMove(Y) = .MoveTutor(X)
                    TempSource(Y) = Temp
                Next X
            End If
            'Now do RBY if in GSC w/Trades
            If GameVersion = nbGSCTrade Then
                For X = 1 To UBound(.RBYMoves)
                    Y = Y + 1
                    TempMove(Y) = .RBYMoves(X)
                    TempSource(Y) = "RBY Level"
                Next X
                For X = 1 To UBound(.RBYTM)
                    Y = Y + 1
                    TempMove(Y) = .RBYTM(X)
                    'Surfing Pikachu - Change HM to Stadium
                    If (.No = 25 Or .No = 26) And TempMove(Y) = 216 Then
                        TempSource(Y) = "Stadium"
                    Else
                        TempSource(Y) = "RBY " & Moves(.RBYTM(X)).OldTM
                    End If
                Next X
            End If
            If GameVersion = nbRBYTrade Or GameVersion = nbGSCTrade Then
                For X = 1 To UBound(.SpecialMoves)
                    Y = Y + 1
                    TempMove(Y) = .SpecialMoves(X)
                    TempSource(Y) = "NYPC"
                    'Moves learned by OTHER than NYPC (not many of 'em)
                    Select Case .No
                        'Amnesia Psyduck in GSC w/Trades
                        Case 54, 55
                            If TempMove(Y) = 6 Then TempSource(Y) = "Stadium"
                        'Baton Pass Farfetch'd
                        Case 83
                            If TempMove(Y) = 12 Then TempSource(Y) = "Stadium 2"
                        'Earthquake Gligar
                        Case 207
                            If TempMove(Y) = 55 Then TempSource(Y) = "Stadium 2"
                        'Extremespeed Dratini
                        Case 147 To 149
                            If TempMove(Y) = 61 Then TempSource(Y) = "Crystal Gift"
                        'Odd Egg Pokes
                        Case 35, 36, 39, 40, 124, 126, 236 To 238
                            If TempMove(Y) = 45 Then TempSource(Y) = "Odd Egg"
                    End Select
                Next X
            ElseIf GameVersion = nbTrueGSC Then
                'Manual adds for True GSC - bypasses the Special field to remove NYPC moves (no Growtheons)
                'Baton Pass Farfetch'd
                If .No = 83 Then
                    ReDim Preserve TempMove(1 To UBound(TempMove) + 1)
                    ReDim Preserve TempSource(1 To UBound(TempSource) + 1)
                    Y = Y + 1
                    TempMove(Y) = 12
                    TempSource(Y) = "Stadium 2"
                'Earthquake Gligar
                ElseIf .No = 207 Then
                    ReDim Preserve TempMove(1 To UBound(TempMove) + 1)
                    ReDim Preserve TempSource(1 To UBound(TempSource) + 1)
                    Y = Y + 1
                    TempMove(Y) = 55
                    TempSource(Y) = "Stadium 2"
                'Dizzy Punchers (Odd Egg)
                ElseIf .No = 239 Or .No = 125 _
                    Or .No = 172 Or .No = 25 Or .No = 26 _
                    Or .No = 173 Or .No = 35 Or .No = 36 _
                    Or .No = 174 Or .No = 39 Or .No = 40 _
                    Or .No = 238 Or .No = 124 _
                    Or .No = 240 Or .No = 126 _
                    Or .No = 236 Or .No = 106 Or .No = 107 Or .No = 237 Then
                    ReDim Preserve TempMove(1 To UBound(TempMove) + 1)
                    ReDim Preserve TempSource(1 To UBound(TempSource) + 1)
                    Y = Y + 1
                    TempMove(Y) = 45
                    TempSource(Y) = "Odd Egg"
                'Extremespeed Dratini
                ElseIf .No = 147 Or .No = 148 Or .No = 149 Then
                    ReDim Preserve TempMove(1 To UBound(TempMove) + 1)
                    ReDim Preserve TempSource(1 To UBound(TempSource) + 1)
                    Y = Y + 1
                    TempMove(Y) = 61
                    TempSource(Y) = "Crystal Gift"
                End If
            End If
        End If
    End With
    Y = UBound(TempMove)
    For X = 1 To Y
        If TempMove(X) > 0 Then
            Select Case GameVersion
            Case nbTrueRBY, nbRBYTrade
                If Not Moves(TempMove(X)).RBYMove Then TempMove(X) = 0
            Case nbTrueGSC, nbGSCTrade
                If Not Moves(TempMove(X)).GSCMove Then TempMove(X) = 0
            Case nbTrueRuSa, nbFullAdvance, nbModAdv
                If Not Moves(TempMove(X)).AdvMove Then TempMove(X) = 0
            End Select
            For Z = X + 1 To Y
                If TempMove(X) = TempMove(Z) Then TempMove(Z) = 0
            Next Z
        End If
    Next X
End Sub
Public Function BinDig(ByVal Max As Long) As Integer
    Dim X As Long
    Select Case Max
    Case 0: BinDig = 0
    Case 1: BinDig = 1
    Case Is < 4: BinDig = 2
    Case Is < 8: BinDig = 3
    Case Is < 16: BinDig = 4
    Case Is < 32: BinDig = 5
    Case Is < 64: BinDig = 6
    Case Is < 128: BinDig = 7
    Case Is < 256: BinDig = 8
    Case Is < 510: BinDig = 9
    Case Is < 1024: BinDig = 10
    Case Else
        X = 10
        Do
            X = X + 1
        Loop Until Max < 2 ^ X
        BinDig = X
    End Select
End Function

Public Function FutureText(txtBox As Control, NewAscii As Integer) As String
    Dim LeftPart As String
    Dim MidPart As String
    Dim RightPart As String
    LeftPart = Left(txtBox.Text, txtBox.SelStart)
    MidPart = Chr$(NewAscii)
    RightPart = Right(txtBox.Text, Len(txtBox.Text) - txtBox.SelStart - txtBox.SelLength)
    FutureText = LeftPart & MidPart & RightPart
End Function

Public Function BattleMatrixEx(ByVal Attacker As Byte, ByVal Defender As Byte, Optional ByVal InRBYMode As Boolean = False) As Single
    Dim X As Single
    X = BattleMatrix(Attacker, Defender)
    If InRBYMode Then
        If Attacker = 12 And Defender = 8 Then X = 2
        If Attacker = 8 And Defender = 12 Then X = 2
        If Attacker = 6 And Defender = 2 Then X = 1
        If Attacker = 12 And Defender = 14 Then X = 1
        If Attacker = 14 And Defender = 11 Then X = 0
    End If
    BattleMatrixEx = X
End Function

Public Function FileHex() As String
    Dim Temp As String
    Do
        Temp = FixedHex(Rnd * 65536, 4)
    Loop Until Dir(SlashPath & "*" & Temp & "*", vbHidden) = ""
    FileHex = Temp
End Function

Public Sub ShowHelpContext(plngContextID As Long)
   Dim hWnd As Long
   Dim lshelpFile As String
   Dim hwndHelp As Long
      
   hWnd = GetDesktopWindow

   lshelpFile = SlashPath & "PokeBattle.chm"
   ' do not replace lshelpFile with a function call
   hwndHelp = HtmlHelp(hWnd, lshelpFile, HH_HELP_CONTEXT, plngContextID)
End Sub
Public Function MakeTeamText(PokeArray() As Pokemon) As String
    Dim Temp As String
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Temp = ""
    For X = 1 To 6
        With PokeArray(X)
            If .No > 0 Then
                Temp = Temp & .Name
                If .No = 201 Then
                    Select Case .UnownLetter
                    Case 26: Temp = Temp & " !"
                    Case 27: Temp = Temp & " ?"
                    Case Else: Temp = Temp & " " & Chr$(.UnownLetter + 65)
                    End Select
                End If
                If .GameVersion <> 0 And .GameVersion <> 5 Then
                    If .Gender <> 0 Then Temp = Temp & " (" & Left$(Gender(.Gender), 1) & ")"
                    Temp = Temp & " @ " & Item(.Item)
                End If
                If .Nickname <> .Name And .Nickname <> "" Then Temp = Temp & " ** " & .Nickname
                If .GameVersion = nbTrueRuSa Or .GameVersion = nbFullAdvance Or .GameVersion = nbModAdv Then
                    If .GameVersion = nbModAdv Then
                        Temp = Temp & vbNewLine & "Trait: " & AttributeText(.ModAttr(.AttNum))
                    Else
                        Temp = Temp & vbNewLine & "Trait: " & AttributeText(.PAtt(.AttNum))
                    End If
                    Z = 0
                    Temp = Temp & vbNewLine & "EVs: "
                    If .EV_HP > 0 Then Temp = Temp & .EV_HP & " HP / ": Z = 1
                    If .EV_Atk > 0 Then Temp = Temp & .EV_Atk & " Atk / ": Z = 1
                    If .EV_Def > 0 Then Temp = Temp & .EV_Def & " Def / ": Z = 1
                    If .EV_Spd > 0 Then Temp = Temp & .EV_Spd & " Spd / ": Z = 1
                    If .EV_SAtk > 0 Then Temp = Temp & .EV_SAtk & " SAtk / ": Z = 1
                    If .EV_SDef > 0 Then Temp = Temp & .EV_SDef & " SDef / ": Z = 1
                    If Z = 1 Then
                        Temp = Left$(Temp, Len(Temp) - 3)
                    Else
                        Temp = Temp & "None"
                    End If
                    Temp = Temp & vbNewLine & Nature(.NatureNum).Name & " Nature ("
                    If .NatureNum Mod 5 = .NatureNum \ 5 Then
                        Temp = Temp & "Neutral)"
                    Else
                        If Nature(.NatureNum).StatChg(1) = "1" Then Temp = Temp & "+Atk, "
                        If Nature(.NatureNum).StatChg(2) = "1" Then Temp = Temp & "+Def, "
                        If Nature(.NatureNum).StatChg(3) = "1" Then Temp = Temp & "+Spd, "
                        If Nature(.NatureNum).StatChg(4) = "1" Then Temp = Temp & "+SAtk, "
                        If Nature(.NatureNum).StatChg(5) = "1" Then Temp = Temp & "+SDef, "
                        If Nature(.NatureNum).StatChg(1) = "-1" Then Temp = Temp & "-Atk)"
                        If Nature(.NatureNum).StatChg(2) = "-1" Then Temp = Temp & "-Def)"
                        If Nature(.NatureNum).StatChg(3) = "-1" Then Temp = Temp & "-Spd)"
                        If Nature(.NatureNum).StatChg(4) = "-1" Then Temp = Temp & "-SAtk)"
                        If Nature(.NatureNum).StatChg(5) = "-1" Then Temp = Temp & "-SDef)"
                    End If
                End If
                For Y = 1 To 4
                    If .Move(Y) > 0 Then
                        Temp = Temp & vbNewLine & "- " & Moves(.Move(Y)).Name
                        If .Move(Y) = 91 Then
                            Temp = Temp & " ["
                            If .GameVersion = nbTrueRuSa Or .GameVersion = nbFullAdvance Or .GameVersion = nbModAdv Then
                                Temp = Temp & Element(HiddenPowerTypeAdv(PKMN(X))) & "]"
                            Else
                                Temp = Temp & Element(HiddenPowerType(PKMN(X).DV_Atk, PKMN(X).DV_Def)) & "]"
                            End If
                        End If
                    End If
                Next Y
                Temp = Temp & vbNewLine & vbNewLine
            End If
        End With
    Next X
    If Temp <> "" Then Temp = Left$(Temp, Len(Temp) - 2)
    MakeTeamText = Temp
End Function

Function CheckFileAuth(ByVal FileName As String, ByVal FileSize As Long, ByVal FileCRC As String) As Boolean
    Dim cStream As New cBinaryFileStream
    Dim cCRC32 As New cCRC32
    Dim lCRC32 As Long
    If Not FileExists(FileName) Then CheckFileAuth = False: Exit Function
    cStream.File = FileName
   
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    
    Debug.Print FileName, cStream.Length, Hex(lCRC32)
   
    If FileSize = cStream.Length And FileCRC = Hex(lCRC32) Then CheckFileAuth = True Else CheckFileAuth = False
End Function

Public Function GetMoveCount(ByRef PKMN As Pokemon, ByVal GameMode As CompatModes)
    Dim AllMoves() As Boolean
    Dim X As Integer
    Dim Y As Integer
    
    ReDim AllMoves(UBound(Moves), 1 To 13)

    Select Case GameMode
        Case nbTrueRBY
            Call ValidMoveArray(PKMN, AllMoves(), nbRBYLevel)
            Call ValidMoveArray(PKMN, AllMoves(), nbRBYTM)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCSpecial)
        Case nbTrueGSC
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCLevel)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCTM)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCEgg)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCTutor)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCSpecial)
        Case nbRBYTrade, nbGSCTrade
            Call ValidMoveArray(PKMN, AllMoves(), nbRBYLevel)
            Call ValidMoveArray(PKMN, AllMoves(), nbRBYTM)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCLevel)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCTM)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCEgg)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCTutor)
            Call ValidMoveArray(PKMN, AllMoves(), nbGSCSpecial)
        Case nbTrueRuSa
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvLevel)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvTM)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvEgg)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvSpecial)
        Case nbFullAdvance, nbModAdv
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvLevel)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvTM)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvEgg)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvSpecial)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvTutor)
            Call ValidMoveArray(PKMN, AllMoves(), nbAdvFL)
        Case Else
            If InVBMode Then Stop
    End Select
    
    For X = 1 To UBound(Moves)
        For Y = 2 To 13
            If AllMoves(X, Y) Then AllMoves(X, 1) = True
        Next
    Next
    Y = 0
    For X = 1 To UBound(Moves)
        If AllMoves(X, 1) Then Y = Y + 1
    Next
    GetMoveCount = Y
End Function

Public Function MD5(TheString As String, Optional AsBytes As Boolean = False) As String
    Dim M As New cMD5
    Dim X As Long
    Dim Y As Long
    Dim Temp As String
    Temp = M.DigestStrToHexStr(TheString)
    Set M = Nothing
    If AsBytes Then
        MD5 = String$(16, vbNullChar)
        For X = 1 To 31 Step 2
            Y = Y + 1
            Mid(MD5, Y, 1) = Chr$(Val("&H" & Mid$(Temp, X, 2)))
        Next X
    Else
        MD5 = Temp
    End If
End Function
Public Function TeamNum(ByVal PokeNum As Byte) As Byte
    Select Case PokeNum
    Case 1, 3: TeamNum = 1
    Case 2, 4: TeamNum = 2
    End Select
End Function


Public Sub ApplyDBMod()
    Dim X As Long
    Dim Text As String
    On Error GoTo ApplyDBMod_Error
    RestoreDB
    DBModHash = MD5(DBModStr, True)

    If Len(DBModStr) = 0 Then Exit Sub
    SetSourceStringASM StrPtr(DBModStr)
    
    Do While GetBitsLeftASM > 2
        Select Case StreamOutASM(2)
        Case 0
            Exit Do
        Case 1
            With BasePKMN(StreamOutASM(9))
                'Debug.Print .Name
                X = UBound(.AdvMoves) + 1
                ReDim Preserve .AdvMoves(X)
                .AdvMoves(X) = StreamOutASM(9)
            End With
        Case 2
            With BasePKMN(StreamOutASM(9))
                'Debug.Print .Name
                X = StreamOutASM(7)
                .ModAttr(StreamOutASM(1)) = X
            End With
        Case 3
            With BasePKMN(StreamOutASM(9))
                'Debug.Print .Name
                X = StreamOutASM(2)
                Text = CStr(StreamOutASM(9))
                For X = 1 To X
                    Text = Text & "+" & CStr(StreamOutASM(9))
                Next X
                .IllegalMod = .IllegalMod & "|" & Text
            End With
        End Select
    Loop
    For X = 1 To 6
        If PKMN(X).No > 0 Then
            With PKMN(X)
                .ModAttr(0) = BasePKMN(.No).ModAttr(0)
                .ModAttr(1) = BasePKMN(.No).ModAttr(1)
                .IllegalMod = BasePKMN(.No).IllegalMod
                .AdvMoves = BasePKMN(.No).AdvMoves
            End With
        End If
    Next X
    SaveSetting "NetBattle", "Recent Files", "Mod", DBModName
    Exit Sub

ApplyDBMod_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ApplyDBMod of Module Code"
    If InVBMode Then Stop: Resume
End Sub
Public Sub RestoreDB()
    Dim X As Long
    For X = 1 To UBound(BasePKMN)
        With BasePKMN(X)
            ReDim Preserve .AdvMoves(0 To .TotalAdvMoves)
            .ModAttr(0) = .PAtt(0)
            .ModAttr(1) = .PAtt(1)
            .IllegalMod = vbNullString
        End With
    Next X
End Sub




Public Sub BenchDLL()
    Dim X As Long
    Dim Y As Long
    Dim C(5) As Long
    Dim T As Single
    Dim Test As String
    Dim Test2 As String
'    Test = "10100010 01010101 00101010 00101001 101000101 01011110"
'    Test = Replace(Test, " ", vbNullString)
'    Test = Bin2Chr(Test)
'
'    SetParseBinStringASM StrPtr(Test)
'    Debug.Print GetParsePosASM
'    Debug.Print streamoutasm(8)
'    Debug.Print GetParsePosASM
'    Debug.Print streamoutasm(5)
'    Debug.Print GetParsePosASM
'    Debug.Print streamoutasm(5)
    
    
    
    Test = "        "
    Randomize Timer
    T = Timer + 0.5
    While T > Timer
        X = Int(Rnd * 256)
        Test = Dec2BinASM(X, 8)
        C(0) = C(0) + 1
    Wend
    
    T = Timer + 0.5
    While T > Timer
        X = Int(Rnd * 256)
        Test = Dec2BinASM(X, 8)
        Y = Bin2Dec(Test)
        C(1) = C(1) + 1
    Wend
    
    T = Timer + 0.5
    While T > Timer
        X = Int(Rnd * 256)
        Test = Dec2BinASM(X, 8)
        Y = Bin2DecASM(StrPtr(Test))
        C(2) = C(2) + 1
    Wend
    Test = "                "
    T = Timer + 0.5
    While T > Timer
        X = Int(Rnd * 65536)
        Test = Dec2BinASM(X, 16)
        C(3) = C(3) + 1
    Wend
    
    T = Timer + 0.5
    While T > Timer
        X = Int(Rnd * 65536)
        Test = Dec2BinASM(X, 16)
        Y = Bin2Dec(Test)
        C(4) = C(4) + 1
    Wend
    
    T = Timer + 0.5
    While T > Timer
        X = Int(Rnd * 65536)
        Test = Dec2BinASM(X, 16)
        Y = Bin2DecASM(StrPtr(Test))
        C(5) = C(5) + 1
    Wend
    
    MsgBox "1Way 8B: " & C(0) & vbNewLine & _
           "VB 8B: " & C(1) & vbNewLine & _
           "Asm 8B: " & C(2) & vbNewLine & _
           "1Way 16B: " & C(3) & vbNewLine & _
           "VB 16B: " & C(4) & vbNewLine & _
           "Asm 16B: " & C(5)

End Sub
Public Sub SaveDBMod()
    Dim FSO As Object
    Dim X As Long
    Dim Path As String
    Dim File As String
    On Error GoTo ETrap
    Path = SlashPath & "Database Mods\"
    File = Replace(DBModName, "<", "_")
    File = Replace(File, ">", "_")
    File = Replace(File, ":", "_")
    File = Replace(File, "?", "_")
    File = Replace(File, """", "_")
    File = Replace(File, "*", "_")
    File = Replace(File, "|", "_")
    File = Replace(File, "/", "_")
    File = Replace(File, "\", "_")
    
    File = Path & File & ".mod"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(Path) Then
        FSO.CreateFolder Path
    End If
    If FileExists(File) Then Kill File
    X = FreeFile
    Open File For Binary Access Write As #X
    Put #X, , DBModStr
    Close #X
ETrap:
End Sub
Public Sub LoadDBMod(File As String)
    Dim X As Long
    Dim Temp As String
    On Error GoTo ETrap
    X = FreeFile
    Open File For Binary Access Read As #X
    DBModStr = String(LOF(X), vbNullChar)
    Get #X, , DBModStr
    Close #X
    Temp = Dir(File)
    Temp = Left$(Temp, InStrRev(Temp, ".") - 1)
    DBModName = Temp
    ApplyDBMod
ETrap:
End Sub
Public Sub SetRedraw(hWnd As Long, RedrawState As Boolean, Optional RefreshNow As Boolean = True)
    If RedrawState Then
        SendMessage hWnd, WM_SETREDRAW, 1, ByVal 0&
        If RefreshNow Then RedrawWindow hWnd, ByVal 0&, 0, RDW_ALLCHILDREN Or RDW_UPDATENOW Or RDW_INVALIDATE
    Else
        SendMessage hWnd, WM_SETREDRAW, 0, ByVal 0&
    End If
End Sub
Public Function FilterIllegalChars(ByVal Text As String, Optional ByVal AllowNewLines As Boolean = False) As String
    Dim X As Long
    For X = 1 To Len(Text)
        Select Case Asc(Mid$(Text, X, 1))
        Case 0 To 9, 11, 12, 15 To 31, 127, 129, 141, 143, 144
            Mid(Text, X, 1) = " "
        Case 10, 13
            If Not AllowNewLines Then Mid(Text, X, 1) = " "
        End Select
    Next X
    FilterIllegalChars = Text
End Function

Public Sub ListViewNumberSort(LView As ListView, ByVal Column As Long)
    Dim X As Long
    Dim C As Long

    With LView
        SetRedraw .hWnd, False
        C = .ListItems.count
        If Column = 1 Then
            For X = 1 To C
                With .ListItems(X)
                    .Tag = Len(.Text)
                    .Text = String$(8 - Len(.Text), "0") & .Text
                End With
            Next X
            .SortKey = 0
            .Sorted = True
            .Sorted = False
            For X = 1 To C
                With .ListItems(X)
                    .Text = Right$(.Text, .Tag)
                End With
            Next X
        Else
            Column = Column - 1
            For X = 1 To C
                With .ListItems(X)
                    .Tag = Len(.SubItems(Column))
                    .SubItems(Column) = String$(8 - Len(.SubItems(Column)), "0") & .SubItems(Column)
                End With
            Next X
            .SortKey = Column
            .Sorted = True
            .Sorted = False
            For X = 1 To C
                With .ListItems(X)
                    .SubItems(Column) = Right$(.SubItems(Column), .Tag)
                End With
            Next X
        End If
        SetRedraw .hWnd, True
    End With
End Sub
Public Function ApplyCSFilter(ByVal Text As String) As String
    Dim X As Long
    ApplyCSFilter = Text
    On Error GoTo ETrap
    For X = 0 To UBound(CSFilter)
        ApplyCSFilter = Replace(ApplyCSFilter, CSFilter(X), String$(Len(CSFilter(X)), "*"))
    Next X
ETrap:
End Function


Public Sub AddNYPCMove(Poke As String, TheMove As String, Optional SkipEvos As Boolean = False)
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim A As Long
    X = GetPokeNum(Poke)
    Y = GetMoveNum(TheMove)
    If X = 0 Or Y = 0 Then
        Stop
    Else
        With BasePKMN(X)
            If UBound(.AdvSpecial) = 1 And .AdvSpecial(1) = 0 Then
                .AdvSpecial(1) = Y
            Else
                For Z = 1 To UBound(.AdvSpecial)
                    If .AdvSpecial(Z) = Y Then Exit Sub
                    If .AdvSpecial(Z) > Y Then Exit For
                Next Z
                ReDim Preserve .AdvSpecial(UBound(.AdvSpecial) + 1)
                For Z = UBound(.AdvSpecial) To Z + 1 Step -1
                    .AdvSpecial(Z) = .AdvSpecial(Z - 1)
                    If .AdvSpecial(Z) = Y Then Exit Sub
                Next Z
                .AdvSpecial(Z) = Y
            End If
            If Not SkipEvos Then
                For Y = 1 To 5
                    If .Evo(Y) <> 0 Then
                        AddNYPCMove BasePKMN(.Evo(Y)).Name, TheMove, True
                    End If
                Next Y
                Beep
            End If
        End With
    End If
End Sub
Public Sub PrintOutput()
    Dim X As Long
    Dim Y As Long
    Dim F As Long
    Dim Temp As String
    F = FreeFile
    Open "C:\output.txt" For Output As #F
    For X = 1 To UBound(BasePKMN)
        Temp = vbNullString
        For Y = 1 To UBound(BasePKMN(X).AdvSpecial)
            If BasePKMN(X).AdvSpecial(Y) <> 0 Then
                Temp = Temp & CStr(BasePKMN(X).AdvSpecial(Y)) & ","
            End If
        Next Y
        Print #F, Temp
    Next X
    Close #F
End Sub
Public Sub DummyFunction()
    Dim Temp As String
    Dim Build As String
    Dim X As Long
    Dim Y() As String
    Dim Z As Long
    Call Dummy2
    Exit Sub
    Open "c:\hex.txt" For Input As #5
    Do Until EOF(5)
        Line Input #5, Temp
        Y = Split(Temp, " ")
        For Z = 0 To UBound(Y)
            If Len(Y(Z)) > 0 Then
                Build = Build & Dec2Bin(CLng("&H" & Y(Z)), 16)
            End If
        Next Z
    Loop
    Close #5
    SaveSetting "NetBattle", "Server", "dbmod", Build
End Sub
Private Sub Dummy2()
    Dim B() As Boolean
    Dim X As Long
    Dim Y As Long
    ReDim B(0 To GetMoveNum("Aerial Ace") - 1)
    For X = 1 To 251
        If X <> 235 Then
        With BasePKMN(X)
            For Y = LBound(.BaseMoves) To UBound(.BaseMoves)
                If .BaseMoves(Y) = GetMoveNum("mimic") Then Stop
                B(.BaseMoves(Y)) = True
            Next Y
        End With
        End If
    Next X
    For X = 1 To UBound(B)
        If Not B(X) Then Debug.Print Moves(X).Name
    Next X
End Sub
