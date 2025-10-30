Attribute VB_Name = "Code"
'--------------------------------------------------------------------------------------
'Pokémon NetBattle
'By "TV's Ian" Murray
'Project begun in mid-2001
'First public release in early 2002 (v0.8.0)
'Project homepage: http://pnb.hubert.us
'--------------------------------------------------------------------------------------
'File type registration from http://www.vbsquare.com/files/tip458.html
'Wave from a resource file from http://www.syix.com/wpsjr1/
'Graphics from a database code from:
'Everything else is either me or MSDN.
'--------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

'Pokemon data
Public Type Pokemon
    No As Integer
    GSNo As Integer
    AdvNo As Integer
    Image As String
    Name As String
    Nickname As String
    Type1 As Integer
    Type2 As Integer
    Attribute As Integer
    Color1 As Integer
    Color2 As Integer
    Move(4) As Integer
    MoveDisabled(4) As Boolean
    MaxPP(4) As Integer
    PP(4) As Integer
    Item As Integer
    Condition As Integer
    ConditionCount As Integer
    BaseHP As Integer
    BaseAttack As Integer
    BaseDefense As Integer
    BaseSpeed As Integer
    BaseSAttack As Integer
    BaseSDefense As Integer
    MaxHP As Integer
    HP As Integer
    HPDV As Integer
    Attack As Integer
    AttackDV As Integer
    Defense As Integer
    DefenseDV As Integer
    Speed As Integer
    SpeedDV As Integer
    SpecialAttack As Integer
    SpecialDefense As Integer
    SpecialDV As Integer
    SpecialRBY As Integer
    Level As Integer
    BaseMoves(251) As Integer
    MachineMoves(251) As Integer
    BreedingMoves(251) As Integer
    RBYMoves(251) As Integer
    SpecialMoves(251) As Integer
    MoveTutor(3) As Integer
    StartsWith As Integer
    PercentFemale As Integer
    Gender As Integer
    'This is the only really strange one - it sets to their position in your lineup.
    'Used for copying info from current to team.
    TeamNumber As Integer
End Type

'Pokedex Text
Public Type PokeDexInfo
    RedBlue As String
    Yellow As String
    Gold As String
    Silver As String
    Ruby As String
    Sapphire As String
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
    StuckCount As Integer
    'Record last 10 moves used
    MoveUsed(10) As Integer
    'Seeded
    LeechSeed As Boolean
    'Attracted
    Attract As Boolean
    'Confused
    Confuse As Boolean
    ConfuseCounter As Integer
    'For Solarbeam, Hyper Beam, etc.
    Charging As Boolean
    Recharging As Boolean
    'Pokemon is out of the playfield
    'May still be hit by certain attacks
    Dig As Boolean
    Fly As Boolean
    'Pokemon can't be switched
    Locked As Boolean
    'Pokemon has successfully used Lock-On or Mind Reader
    LockOn As Boolean
    'Pokemon has used Foresight - Normal or Fighting can hit Ghosts
    Foresight As Boolean
    'These two work with Bide
    BideDamage As Integer
    BideCount As Integer
    'For hit 2-5 turn moves - Fire Spin, Whirlpool, etc.
    RepeatMove As Integer
    RepeatCount As Integer
    'Pokemon has used Destiny Bond
    DestinyBond As Boolean
    'Counts up Toxic's damage
    ToxicCount As Single
    'Perish Song Counter
    PerishSong As Integer
    'If the opponent is hit with Nightmare
    Nightmare As Boolean
    'For Counter/Mirror Coat
    LastDamage As Integer
    LastSDamage As Integer
    'For Rollout
    DefenseCurl As Boolean
    Rollout As Integer
    FuryCutter As Integer
    Curse As Boolean
    'For the Minimize/Stomp combo
    Minimize As Boolean
    'Disable
    DisabledMove As Integer
    DisableCount As Integer
    'Remaining HP for a Substitute
    Substitute As Integer
    'For the three Protect moves
    ProtectPercent As Integer
    'Under the effects of Mist
    Mist As Boolean
    'Encore
    Encore As Boolean
    EncoreMove As Integer
    EncoreReplaced As Integer
    EncoreDuration As Integer
    MimicedMove As Integer
    RageCounter As Integer
End Type

'Things that don't change between switches.
Public Type TeamCond
    SafeGuardCount As Integer
    ReflectCount As Integer
    LightScreenCount As Integer
    Spikes As Boolean
End Type

'Move stuff
Public Type Move
    'Need to use this one for IconList funkiness
    ID As Integer
    Name As String
    Type As Integer
    Power As Integer
    Accuracy As Integer
    PP As Integer
    'Text is a description - it comes up as a tooltip on the team builder
    'Note to self - add it as a tooltip on the battle screen
    Text As String
    SpecialPercent As Integer
    SpecialEffect As Integer
    WorksRight As Boolean
    BrightPowder As Boolean
    KingsRock As Boolean
    OldMove As Boolean
    SelfMove As Boolean
    OldTM As String
    NewTM As String
    SubstituteBlocks As Boolean
End Type

'Trainer info
Public Type Trainer
    Picture As Integer
    Name As String
    'Decides which set of graphics to use for your own Pokemon
    Version As Integer
    ProgVersion As String
    'Extra comes up as a tooltip on the battle screen, and in the challenge window.
    Extra As String
    WinMess As String
    LoseMess As String
End Type

'Future Sight attacks
Public Type FutureSightData
    Turn As Integer
    Team As Integer
    AttackPower As Integer
    CHit As Boolean
    Hit As Boolean
End Type

Public Type ServerPKMNData
    Item As Integer
    Level As Integer
    Gender As Integer
    Nickname As String
    HPDV As Integer
    AttackDV As Integer
    DefenseDV As Integer
    SpeedDV As Integer
    SpecialDV As Integer
    Move(4) As Integer
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
    TeamString As String
    PokeData(6) As ServerPKMNData
    SkipXOR As Boolean
    PKMN(6) As Integer
    PKMNImage(6) As String
    BattlingWith As Integer
    Rank As String
    RBYCompatible As Boolean
    Wins As Long
    Losses As Long
    Ties As Long
    Disconnect As Long
    Unrated As Boolean
    PingTime As Single
    Speed As String
    Ignore() As Integer
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

'Main database connection
Global PokeData As New ADODB.Connection

'Debug variable
'Double-click on your Pokemon to toggle during battle
'Type a capital D on the loader window to activate there.
Global DebugMode As Boolean

'BasePKMN = default values for Max Gene Pokemon
Global BasePKMN(386) As Pokemon

'Your team, unaffected by battles or anything else.
Global StoredPKMN(6) As Pokemon

'Your team and your opponent's team, as used in battle
Global PKMN(6) As Pokemon
Global EnemyPKMN(6) As Pokemon

'Swap space for Expert Mode.
'Has to do a Global, because VB doesn't let you do Public user types in a form.
Global ExpertPKMN As Pokemon

'Move information
Global Moves(400) As Move

'Type effectiveness chart - (AttackType,DefendType)
Global BattleMatrix(17, 17) As Single

'Trainer info
Global You As Trainer
Global StationID As String

'To transfer info between the registry, network dialog, and battle screen
'IsServer also determines a few things during battle
Global ServerAddress As String
Global IsServer As Boolean

'Strings for display purposes
Global Gender(2) As String
Global Weather(3) As String
Global Condition(8) As String
Global Element(0 To 17) As String
Global Item(100) As String
Global RuleText(9) As String
Global FlavorText(1000) As String
Global AttributeText(77) As String
Global ColorText(10) As String
Global Stat(10) As String
Global PokedexText(386) As PokeDexInfo

'Server stuff, mostly
'Anything that didn't need to be Global is in the forms' declarations
'Might still be able to clean it up, but it might not be worth it.
Global RelayServer As Boolean
Global Player(256) As MSPlayer
Global IsLoaded(256) As Boolean
Global Chances(256) As Integer
Global Disconnecting(256) As Boolean
Global RuleSelected(100) As Integer
Global SelectedPlayer As Integer
Global Challenge As Boolean
Global YourNumber As Integer
Global GameType As Integer
Global ChallengeNumber As Integer
Global ChallengePending As Boolean
Global ICalled As Boolean
Global Ranking As String
Global Battling As Boolean
Global ListenWrong As Boolean

'Options
Global RecentFiles(4) As String
Global SoundOption As Integer
Global MusicOption As Integer
Global AnimOption As Integer
Global AutoScan As Integer
Global AskOnUpdate As Integer
Global BMessStyle As Integer
Global AllowViewing As Integer
Global LogPrompt As Integer
Global SavedPassword As String
Global FancyText As Boolean
Global SoundFile(20) As String
Global RecentServer(100) As String
Global UseAI As Integer
Global LFile(255) As LPlug
Global CurrLang As String
Global GetSpeed As Integer
Global TBSort As Integer

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

'For internal use only.
Global DoneLoading As Boolean
Global DBChecksum As String
Global PasswordBoxTitle As String
Global PasswordBoxCaption As String
Global WinDir As String
Global SysDir As String
Global DidScan As Boolean
Global DataPath As String
Global SlashPath As String
Global DBPassword As String
Global RBYCompatible As Boolean

'My constants - networking and ranking info
Global Const NetChunkSize = 256
Global Const BaseURL = "http://www.tvsian.com/netbattle/nbupdate/"
Global Const LowestRank = -593
Global Const HighestRank = 13495
Global Const RegAddress = "masamune.zapto.org" 'EDIT
'--------------------------------------------------------------------------------------
'Stuff below here comes from sample code from the controls, or code from other sources.
'--------------------------------------------------------------------------------------

'These are for the load-graphics-from-a-database code
Public objDB As Database
Public gsFileName As String
Public gsDrive As String
Public gsPath As String
Public gErrFormName As String
Public Const conChunkSize = 8192

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
Public Declare Function GetWindowsDirectoryA Lib "KERNEL32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Get the system folder
Public Declare Function GetSystemDirectoryA Lib "KERNEL32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Flash the taskbar item
Public Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'For generating the station ID
Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Sub Main()
    Dim TempVar As String
    Dim sPath As String
    Dim sBuf As String
    Dim cSize As Long
    Dim retval As Long
    
    DoneLoading = False
    RunningServer = False
    
    Randomize Timer 'EDIT
    
    'Set up the main database connection
    If Right$(App.Path, 1) <> "\" Then
        SlashPath = App.Path & "\"
    Else
        SlashPath = App.Path
    End If
    DataPath = SlashPath & "PokeDB.mdb"

    DidScan = False
    'Get the Windows folder name
    sBuf = String(255, 0)
    cSize = 255
    retval = GetWindowsDirectoryA(sBuf, cSize)
    sBuf = Left(sBuf, retval)
    WinDir = sBuf
    If Right(WinDir, 1) <> "\" Then WinDir = WinDir & "\"
    
    'Get the system folder name
    sBuf = String(255, 0)
    cSize = 255
    retval = GetSystemDirectoryA(sBuf, cSize)
    sBuf = Left(sBuf, retval)
    SysDir = sBuf
    If Right(SysDir, 1) <> "\" Then SysDir = SysDir & "\"
    
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
    Condition(1) = "Normal"
    Condition(2) = "Poisoned"
    
    'Note: 3 is actually Toxic, it remins separate due to Baton Pass weirdness.
    Condition(3) = "Poisoned"
    Condition(4) = "Sleeping"
    Condition(5) = "Burned"
    Condition(6) = "Paralyzed"
    Condition(7) = "Frozen"
    Condition(8) = "Fainted"
    
    'Item names
    Item(0) = "(None)"
    Item(1) = "Berry"
    Item(2) = "Berry Juice"
    Item(3) = "Bitter Berry"
    Item(4) = "Burnt Berry"
    Item(5) = "Gold Berry"
    Item(6) = "Ice Berry"
    Item(7) = "Mint Berry"
    Item(8) = "Miracle Berry"
    Item(9) = "Mystery Berry"
    Item(10) = "Paralyzecure Berry"
    Item(11) = "Poisoncure Berry"
    Item(12) = "Berserk Gene"
    Item(13) = "Black Belt"
    Item(14) = "Black Glasses"
    Item(15) = "Bright Powder"
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
    Item(29) = "Nevermelt Ice"
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
    Item(41) = "Twisted Spoon"
    
    'Weather Conditions
    Weather(0) = "Normal"
    Weather(1) = "Rainy"
    Weather(2) = "Sunny"
    Weather(3) = "Sandstorm"
    
    'Rule Text
    RuleText(1) = "Sleep/Freeze Clause"
    RuleText(2) = "Self-KO Clause"
    RuleText(3) = "Enable One-Hit KOs"
    RuleText(4) = "Apply PP Ups"
    RuleText(5) = "Stadium Mode"
    RuleText(6) = "R/B/Y Mode"
    RuleText(7) = "One Legendary per Team"
    RuleText(8) = "No Legendaries"
    RuleText(9) = "Use Crystal/Stadium Present"
    
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
    AttributeText(2) = "In the Rain"
    AttributeText(3) = "Acceleration"
    AttributeText(4) = "Kabuto Armor"
    AttributeText(5) = "Solid"
    AttributeText(6) = "Moisture"
    AttributeText(7) = "Flexible"
    AttributeText(8) = "Sand Fall"
    AttributeText(9) = "Static Electricity"
    AttributeText(10) = "Storage of Electricity"
    AttributeText(11) = "Storage of Water"
    AttributeText(12) = "Stolidity"
    AttributeText(13) = "No Weather"
    AttributeText(14) = "Compound Eyes"
    AttributeText(15) = "Insomnia"
    AttributeText(16) = "Color Change"
    AttributeText(17) = "Immunity"
    AttributeText(18) = "Moraibi"
    AttributeText(19) = "Moth Scales"
    AttributeText(20) = "My Pace"
    AttributeText(21) = "Sucker"
    AttributeText(22) = "Menace"
    AttributeText(23) = "Kagefumi"
    AttributeText(24) = "Fishskin"
    AttributeText(25) = "Mysterious Protection"
    AttributeText(26) = "Floating"
    AttributeText(27) = "Emission"
    AttributeText(28) = "Synchronize"
    AttributeText(29) = "Clear Body"
    AttributeText(30) = "Healing Nature"
    AttributeText(31) = "Lightning Rod"
    AttributeText(32) = "Angel Blessing"
    AttributeText(33) = "Swim (Smoothly)"
    AttributeText(34) = "Chlorophyl"
    AttributeText(35) = "Misfortune"
    AttributeText(36) = "Trance"
    AttributeText(37) = "Charge Power"
    AttributeText(38) = "Poison Spike"
    AttributeText(39) = "Mind Power"
    AttributeText(40) = "Magma Armor"
    AttributeText(41) = "Water Veil"
    AttributeText(42) = "Magnetism"
    AttributeText(43) = "Ingratitude"
    AttributeText(44) = "Rain Saucer"
    AttributeText(45) = "Sanaokoshi"
    AttributeText(46) = "Pressure"
    AttributeText(47) = "Deep Desire"
    AttributeText(48) = "Early Rising"
    AttributeText(49) = "Body of Fire"
    AttributeText(50) = "No Run"
    AttributeText(51) = "Sharp Eye"
    AttributeText(52) = "Superhuman Strength Basami"
    AttributeText(53) = "Gather Things"
    AttributeText(54) = "Lazy (Person)"
    AttributeText(55) = "Needle Drill"
    AttributeText(56) = "ATTRACTive Body"
    AttributeText(57) = "Plus"
    AttributeText(58) = "Minus"
    AttributeText(59) = "Moody Person"
    AttributeText(60) = "Adhesion"
    AttributeText(61) = "Shedding"
    AttributeText(62) = "Will-Power"
    AttributeText(63) = "Mysterious (Fish) Scale"
    AttributeText(64) = "Liquid Sludge"
    AttributeText(65) = "Fresh Verdure"
    AttributeText(66) = "Raging Fire"
    AttributeText(67) = "Raging Stream"
    AttributeText(68) = "Bug's Notice"
    AttributeText(69) = "Rock Head"
    AttributeText(70) = "Drought"
    AttributeText(71) = "Doodlebug"
    AttributeText(72) = "Yaruki"
    AttributeText(73) = "White Smoke"
    AttributeText(74) = "Yoga Power"
    AttributeText(75) = "Shell Armor"
    AttributeText(76) = "Noise"
    AttributeText(77) = "Air-Lock"
    
    'Stats
    Stat(1) = "HP"
    Stat(2) = "Attack"
    Stat(3) = "Defense"
    Stat(4) = "Speed"
    Stat(5) = "Special"
    Stat(6) = "Special Attack"
    Stat(7) = "Special Defense"
    Stat(8) = "Accuracy"
    Stat(9) = "Evasion"
    
    'Flavor (Battle) Text
    FlavorText(1) = "%1, go!"
    FlavorText(2) = "%1, return!"
    FlavorText(3) = "%1 fainted!"
    FlavorText(4) = "%1 used %2"
    FlavorText(5) = "%1 missed!"
    FlavorText(6) = "%1 has been poisoned!"
    FlavorText(7) = "%1 has been badly poisoned!"
    FlavorText(8) = "%1 fell asleep!"
    FlavorText(9) = "%1 has been burned!"
    FlavorText(10) = "%1 has been paralyzed!"
    FlavorText(11) = "%1 has been frozen!"
    FlavorText(12) = "%1 is hurt by the poison!"
    FlavorText(13) = "%1 is still asleep!"
    FlavorText(14) = "%1 woke up!"
    FlavorText(15) = "%1 is hurt by the burn!"
    FlavorText(16) = "%1 is fully paralyzed!"
    FlavorText(17) = "%1 is frozen solid!"
    FlavorText(18) = "%1 thawed out!"
    FlavorText(19) = "%1 healed using a %2"
    FlavorText(20) = "%1 healed using %2"
    FlavorText(21) = "%1 is hurt by the Curse!"
    FlavorText(22) = "%1 is still having a Nightmare!"
    FlavorText(23) = "%1 is doing an Encore!"
    FlavorText(24) = "%1 finished the Encore!"
    FlavorText(25) = "%1 is hurt by the Sandstorm!"
    FlavorText(26) = "The sun continues to shine!"
    FlavorText(27) = "The rain is continuing to fall!"
    FlavorText(28) = "The sandstorm continues!"
    FlavorText(29) = "The sun stopped shining!"
    FlavorText(30) = "The rain stopped!"
    FlavorText(31) = "The sandstorm ended!"
    FlavorText(32) = "%1 became confused!"
    FlavorText(33) = "%1 is confused!"
    FlavorText(34) = "It hurt itself in it's confusion!"
    FlavorText(35) = "%1 is attracted to %2!"
    FlavorText(36) = "%1's attraction kept it from attacking!"
    FlavorText(37) = "%1 recovered health!"
    FlavorText(38) = "%1's %2 rose!"
    FlavorText(39) = "%1's %2 rose greatly!"
    FlavorText(40) = "%1's %2 fell!"
    FlavorText(41) = "%1's %2 fell greatly!"
    FlavorText(42) = "%1 sucked health from %2"
    FlavorText(43) = "Critical Hit!"
    FlavorText(44) = "Hit %1 times!"
    FlavorText(45) = "%1 backed off!"
    FlavorText(46) = "%1 flinched!"
    FlavorText(47) = "%1's rage is building!"
    FlavorText(48) = "%1 transformed into %2!"
    FlavorText(49) = "%1's turn!"
    FlavorText(50) = "%1 is haunted by nightmares!"
    FlavorText(51) = "%1 won!"
    FlavorText(52) = "%1 lost!"
    FlavorText(53) = "Tie game!"
    FlavorText(54) = "%1 is taking in sunlight!"
    FlavorText(55) = "%1 is recharging!"
    FlavorText(56) = "%1 is charging!"
    FlavorText(57) = "%1 has locked on!"
    FlavorText(58) = "%1 can't escape!"
    FlavorText(59) = "%1 is hurt by spikes!"
    FlavorText(60) = "Coins scattered everywhere!"
    FlavorText(61) = "The move failed!"
    FlavorText(62) = "%1 has been seeded!"
    FlavorText(63) = "%1 converted its type to %2"
    FlavorText(64) = "%1's %2 activated!"
    FlavorText(65) = "The move is disabled!"
    FlavorText(66) = "Missed!%n%1 crashed!"
    FlavorText(67) = "%1 dug underground!"
    FlavorText(68) = "%1 flew up high!"
    FlavorText(69) = "%1 lowered its head!"
    
    'Build a string out of the version
    You.ProgVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    'Load the recent files
    RecentFiles(1) = GetSetting("NetBattle", "Recent Files", "1", "")
    RecentFiles(2) = GetSetting("NetBattle", "Recent Files", "2", "")
    RecentFiles(3) = GetSetting("NetBattle", "Recent Files", "3", "")
    RecentFiles(4) = GetSetting("NetBattle", "Recent Files", "4", "")
    
    'This is only used when a team is hidden
    BasePKMN(0).Name = "???"
    
    'This bit comes from the load-graphics-from-a-database code
    sPath = App.Path
    Set objDB = OpenDatabase(sPath & "\Graphics.mdb", False)
    'End someone else's code
    
    PokeData.Provider = "Microsoft.Jet.OLEDB.4.0"
    PokeData.Properties("Data Source") = DataPath
    PokeData.Properties("Jet OLEDB:Database Password") = "ginyu4ce"
    PokeData.Open
    
    'Get hard drive serial number (Station ID)
    StationID = Str(GetSerialNumber(Left(WinDir, 3)))
    
    'MainContainer contains some shared controls and code (to reduce overall filesize).
    'They include imageLists, the graphics loader, and the file dialog.
    'When MainContainer unloads, the program ends.
    
    'MainContainer.Show
End Sub

Public Function GetDBChecksum() As String
    'Grabs a total on all the DB entries.
    'For some reason, it doesn't always seem to sync up, so it's
    'not actually checked right now.
    Dim BigDBThingy As Double
    Dim X As Integer
    Dim Y As Integer
    
    BigDBThingy = 0
    For X = 1 To 251
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

Public Function EncryptString(ByVal EncryptMe As String) As String
    'Really stupid one-way string encryption
    'Used for the user passwords.
    'Yes, you can theoretically have more than one string get the same result,
    'But for the minor use it has, it's not important.
    Dim TempVar As Long
    Dim TempVar2 As String
    Dim X As Integer
    
    For X = 1 To Len(EncryptMe)
        TempVar = TempVar + (Asc(Mid(EncryptMe, X, 1)) * X)
    Next
    If Len(EncryptMe) < 10 Then
        TempVar2 = "0" & Len(EncryptMe) & TempVar
    Else
        TempVar2 = Len(EncryptMe) & TempVar
    End If
    EncryptString = TempVar2
End Function

Public Function IsVersionAt(ByVal PVersion As String, ByVal Major As Integer, ByVal Minor As Integer, ByVal Rev As Integer) As Boolean
    'Version checking code.
    'Returns True if the version passed to it (usually an online player) is >= the requested
    Dim MajorVersion As Integer
    Dim MinorVersion As Integer
    Dim Revision As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    
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
    
    Total = (GetHP(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseHP, TempPKMN.HPDV) / 1.5) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseAttack, TempPKMN.AttackDV) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseDefense, TempPKMN.DefenseDV) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseSAttack, TempPKMN.SpecialDV) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseSDefense, TempPKMN.SpecialDV) _
        + GetStat(TempPKMN.Level, BasePKMN(TempPKMN.No).BaseSpeed, TempPKMN.SpeedDV)
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
        Case 150, 249, 250
            Total = Total + (TempPKMN.Level * 5)
        'Legendary birds & dogs, Mew, Celebi
        Case 144 To 146, 151, 243 To 246, 251
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
    Dim X As Integer
    Dim Y As Integer
    Dim MatrixAdjust As Long
    Dim BattleDamage As Single
    Dim LoVal As Long
    Dim HiVal As Long
    
    If Right(App.Path, 1) <> "\" Then
        Open App.Path & "\ranks.csv" For Output As #1
    Else
        Open App.Path & "ranks.csv" For Output As #1
    End If
    Write #1, "Name", "Low", "High"
    For X = 1 To 251
        LoVal = (GetHP(1, BasePKMN(X).BaseHP, 0) / 1.5) _
            + GetStat(1, BasePKMN(X).BaseAttack, 0) _
            + GetStat(1, BasePKMN(X).BaseDefense, 0) _
            + GetStat(1, BasePKMN(X).BaseSAttack, 0) _
            + GetStat(1, BasePKMN(X).BaseSDefense, 0) _
            + GetStat(1, BasePKMN(X).BaseSpeed, 0)
        HiVal = (GetHP(100, BasePKMN(X).BaseHP, 15) / 1.5) _
            + GetStat(100, BasePKMN(X).BaseAttack, 15) _
            + GetStat(100, BasePKMN(X).BaseDefense, 15) _
            + GetStat(100, BasePKMN(X).BaseSAttack, 15) _
            + GetStat(100, BasePKMN(X).BaseSDefense, 15) _
            + GetStat(100, BasePKMN(X).BaseSpeed, 15)
        MatrixAdjust = 0
        For Y = 1 To 17
            If BasePKMN(X).Type2 = 0 Then
                BattleDamage = BattleMatrix(Y, BasePKMN(X).Type1)
            Else
                BattleDamage = BattleMatrix(Y, BasePKMN(X).Type1) * BattleMatrix(Y, BasePKMN(X).Type2)
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
        LoVal = LoVal + MatrixAdjust
        HiVal = HiVal + MatrixAdjust
        'Skew for overly powerful PKMN
        Select Case X
            'Mewtwo, Lugia, Ho-Oh
            Case 150, 249, 250
                HiVal = HiVal + 500
                LoVal = LoVal + 5
            'Legendary birds & dogs, Mew, Celebi
            Case 144 To 146, 151, 243 To 246, 251
                HiVal = HiVal + 250
                LoVal = LoVal + 2.5
            'Snorlax, Dragonite, Tyranitar
            Case 143, 149, 248
                HiVal = HiVal + 125
                LoVal = LoVal + 1.25
        End Select
        Write #1, BasePKMN(X).Name, LoVal, HiVal
    Next
    Close #1
    MsgBox "Individual totals written to ranks.csv", vbInformation, "Done"
End Sub

Public Function FileExists(ByVal FileName As String) As Boolean
    'Determines if a file exists
    'Used by the auto-updater dowwnloader thingy.
    On Error GoTo NoFile
    Open FileName For Input As #1
    Close #1
    FileExists = True
    Exit Function
NoFile:
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
    Dim X  As Integer
    Dim Temp As String
     
    If Not FileExists(SlashPath & "servlist.txt") Then
        Temp = ""
        Open SlashPath & "servlist.txt" For Output As #1
        Write #1, "hubert.dnsalias.net"
        For X = 1 To 99
            Write #1, Temp
        Next
        Close #1
    End If
    
    Open SlashPath & "servlist.txt" For Input As #1
    For X = 1 To 100
        Input #1, RecentServer(X)
    Next
    Close #1
End Sub

Public Sub UpdateServerList(ByVal LastServer As String)
    Dim X As Integer
    Dim CurrentPosition As Integer
    
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
    Open SlashPath & "servlist.txt" For Output As #1
    For X = 1 To 100
        Write #1, RecentServer(X)
    Next
    Close #1
End Sub

Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
End Function

Public Function GetStat(ByVal Level As Integer, ByVal Base As Integer, ByVal DV As Integer) As Integer
    GetStat = Int(Level * (Base + DV + 31.9) / 50) + 5
End Function

Public Function GetHP(ByVal Level As Integer, ByVal Base As Integer, ByVal DV As Integer) As Integer
    GetHP = Int(Level * (Base + DV + 31.9 + 50) / 50) + 10
End Function

Public Function ShinyDV(ByVal AttackDV As Integer) As Boolean
    Select Case AttackDV
        Case 15, 14, 11, 10, 7, 6, 3, 2
            ShinyDV = True
        Case Else
            ShinyDV = False
    End Select
End Function

Public Function StringCount(BaseString As String, CountString As String) As Integer
    Dim C As Integer
    Dim P As Integer
    
    C = 0
    P = 1
    While P > 0
        P = InStr(P, BaseString, CountString)
        If P > 0 Then
            P = P + 1
            C = C + 1
        End If
    Wend
    StringCount = C
End Function

Public Function ChooseImage(ByVal Number As Integer, ByVal AtkDV As Integer, ByVal DefDV As Integer, ByVal SpdDV As Integer, ByVal SplDV As Integer, ByVal Ver As Integer)
    Dim AttackVar As Integer
    Dim DefenseVar As Integer
    Dim SpeedVar As Integer
    Dim SpecialVar As Integer
    Dim Unown As Integer
    Dim GFXFile As String
    
    If Number <> 201 Then
        'Decide which graphic to use
        If Ver = 0 Then
            If DefDV = 10 And SpdDV = 10 And SplDV = 10 Then
                GFXFile = Str(Number) + "gs.gif"
            Else
                GFXFile = Str(Number) + "g.gif"
            End If
        Else
            If ShinyDV(AtkDV) And DefDV = 10 And SpdDV = 10 And SplDV = 10 Then
                GFXFile = Str(Number) + "ss.gif"
            Else
                GFXFile = Str(Number) + "s.gif"
            End If
        End If
        GFXFile = Right$(GFXFile, Len(GFXFile) - 1)
    Else
        Select Case AtkDV
            Case 0
                AttackVar = 0
            Case 1
                AttackVar = 0
            Case 2
                AttackVar = 64
            Case 3
                AttackVar = 64
            Case 4
                AttackVar = 128
            Case 5
                AttackVar = 128
            Case 6
                AttackVar = 192
            Case 7
                AttackVar = 192
            Case 8
                AttackVar = 0
            Case 9
                AttackVar = 0
            Case 10
                AttackVar = 64
            Case 11
                AttackVar = 64
            Case 12
                AttackVar = 128
            Case 13
                AttackVar = 128
            Case 14
                AttackVar = 192
            Case 15
                AttackVar = 192
        End Select
        Select Case DefDV
            Case 0
                DefenseVar = 0
            Case 1
                DefenseVar = 0
            Case 2
                DefenseVar = 16
            Case 3
                DefenseVar = 16
            Case 4
                DefenseVar = 32
            Case 5
                DefenseVar = 32
            Case 6
                DefenseVar = 48
            Case 7
                DefenseVar = 48
            Case 8
                DefenseVar = 0
            Case 9
                DefenseVar = 0
            Case 10
                DefenseVar = 16
            Case 11
                DefenseVar = 16
            Case 12
                DefenseVar = 32
            Case 13
                DefenseVar = 32
            Case 14
                DefenseVar = 48
            Case 15
                DefenseVar = 48
        End Select
        Select Case SpdDV
            Case 0
                SpeedVar = 0
            Case 1
                SpeedVar = 0
            Case 2
                SpeedVar = 4
            Case 3
                SpeedVar = 4
            Case 4
                SpeedVar = 8
            Case 5
                SpeedVar = 8
            Case 6
                SpeedVar = 12
            Case 7
                SpeedVar = 12
            Case 8
                SpeedVar = 0
            Case 9
                SpeedVar = 0
            Case 10
                SpeedVar = 4
            Case 11
                SpeedVar = 4
            Case 12
                SpeedVar = 8
            Case 13
                SpeedVar = 8
            Case 14
                SpeedVar = 12
            Case 15
                SpeedVar = 12
        End Select
        Select Case SplDV
            Case 0
                SpecialVar = 0
            Case 1
                SpecialVar = 0
            Case 2
                SpecialVar = 1
            Case 3
                SpecialVar = 1
            Case 4
                SpecialVar = 2
            Case 5
                SpecialVar = 2
            Case 6
                SpecialVar = 3
            Case 7
                SpecialVar = 3
            Case 8
                SpecialVar = 0
            Case 9
                SpecialVar = 0
            Case 10
                SpecialVar = 1
            Case 11
                SpecialVar = 1
            Case 12
                SpecialVar = 2
            Case 13
                SpecialVar = 2
            Case 14
                SpecialVar = 3
            Case 15
                SpecialVar = 3
        End Select
        Unown = Int((AttackVar + DefenseVar + SpeedVar + SpecialVar) / 10 + 1)
        If Unown < 1 Then Unown = Unown + 26
        If Unown > 26 Then Unown = Unown - 26
        If Unown < 10 Then
            GFXFile = "2010" & Right(Str(Unown), Len(Str(Unown)) - 1) & ".gif"
        Else
            GFXFile = "201" & Right(Str(Unown), Len(Str(Unown)) - 1) & ".gif"
        End If
        If DefDV = 10 And SpdDV = 10 And SplDV = 10 Then
            GFXFile = "201s.gif"
        End If
    End If
    ChooseImage = GFXFile
End Function

Public Function FText(ByVal StringNum As Integer, _
    Optional ByVal Param1 As String = "", _
    Optional ByVal Param2 As String = "") As String
    
    Dim Temp As String
    Dim X As Integer
    
    X = InStr(1, LCase(Temp), "%n")
    If X > 0 Then
        Temp = Left(Temp, X - 1) & vbCrLf & Right(Temp, Len(Temp) - (X + 1))
    End If
    
    Temp = FlavorText(StringNum)
    X = InStr(1, Temp, "%1")
    
    If Param1 = "" Or X = 0 Then
        FText = Temp
        Exit Function
    End If
    
    Temp = Left(Temp, X - 1) & Param1 & Right(Temp, Len(Temp) - (X + 1))
    X = InStr(1, Temp, "%1")
    
    If Param2 = "" Or X = 0 Then
        FText = Temp
        Exit Function
    End If
    
    Temp = Left(Temp, X - 1) & Param2 & Right(Temp, Len(Temp) - (X + 1))
    FText = Temp
End Function

Public Function Cap(ByVal Value As Integer) As Integer
    If Value > 999 Then Cap = 999 Else Cap = CInt(Value)
End Function

Public Function Rollover(ByVal Value As Integer, ByVal StadiumMode As Boolean) As Integer
    If StadiumMode Then
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

Public Function Dec(ByVal HexNum As String) As Integer
    Dec = Val("&H" & HexNum)
End Function

Sub LoadPKMNData()
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
    
    'MainContainer.MousePointer = vbHourglass
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Pokemon WHERE Number > 0 ORDER BY Number ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    While Not QueryResults.EOF
        CurrentRecord = QueryResults("Number")
        BasePKMN(CurrentRecord).No = CurrentRecord
        BasePKMN(CurrentRecord).Name = QueryResults("Name")
        BasePKMN(CurrentRecord).Type1 = QueryResults("Type1")
        If Len(QueryResults("Type2")) > 0 Then
            BasePKMN(CurrentRecord).Type2 = QueryResults("Type2")
        End If
        BasePKMN(CurrentRecord).BaseHP = QueryResults("HP")
        BasePKMN(CurrentRecord).BaseAttack = QueryResults("Attack")
        BasePKMN(CurrentRecord).BaseDefense = QueryResults("Defense")
        BasePKMN(CurrentRecord).BaseSpeed = QueryResults("Speed")
        BasePKMN(CurrentRecord).BaseSAttack = QueryResults("SpecialAttack")
        BasePKMN(CurrentRecord).BaseSDefense = QueryResults("SpecialDefense")
        BasePKMN(CurrentRecord).SpecialRBY = QueryResults("SpecialRBY")
        BasePKMN(CurrentRecord).StartsWith = QueryResults("BornWith")
        BasePKMN(CurrentRecord).PercentFemale = QueryResults("Percent Female")
        If Len(QueryResults("Moves")) > 0 Then
            Temp = QueryResults("Moves")
            Y = 1
            P1 = 0
            P2 = InStr(1, Temp, ",")
            While P2 > 0
                BasePKMN(CurrentRecord).BaseMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
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
                BasePKMN(CurrentRecord).MachineMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
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
                BasePKMN(CurrentRecord).BreedingMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
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
                BasePKMN(CurrentRecord).RBYMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
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
                BasePKMN(CurrentRecord).SpecialMoves(Y) = Val(Mid(Temp, P1 + 1, P2 - P1 - 1))
                P1 = P2
                P2 = InStr(P1 + 1, Temp, ",")
                Y = Y + 1
            Wend
        End If
        If Len(QueryResults("Move Tutor")) > 0 Then
            MTTemp = QueryResults("Move Tutor")
            If MTTemp - 4 >= 0 Then
                BasePKMN(CurrentRecord).MoveTutor(1) = 70
                MTTemp = MTTemp - 4
            End If
            If MTTemp - 2 >= 0 Then
                BasePKMN(CurrentRecord).MoveTutor(2) = 98
                MTTemp = MTTemp - 2
            End If
            If MTTemp - 1 >= 0 Then
                BasePKMN(CurrentRecord).MoveTutor(3) = 232
                MTTemp = MTTemp - 1
            End If
        End If
        'Fill in L100 stats for the Pokedex
        BasePKMN(CurrentRecord).Attack = GetStat(100, BasePKMN(CurrentRecord).BaseAttack, 15)
        BasePKMN(CurrentRecord).Defense = GetStat(100, BasePKMN(CurrentRecord).BaseDefense, 15)
        BasePKMN(CurrentRecord).Speed = GetStat(100, BasePKMN(CurrentRecord).BaseSpeed, 15)
        BasePKMN(CurrentRecord).SpecialAttack = GetStat(100, BasePKMN(CurrentRecord).BaseSAttack, 15)
        BasePKMN(CurrentRecord).SpecialDefense = GetStat(100, BasePKMN(CurrentRecord).BaseSDefense, 15)
        BasePKMN(CurrentRecord).MaxHP = GetHP(100, BasePKMN(CurrentRecord).BaseHP, 15)
        Call ScanForDuplicates(CurrentRecord)
        QueryResults.MoveNext
    Wend
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Moves WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    While Not QueryResults.EOF
        CurrentRecord = QueryResults("ID")
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
        Moves(CurrentRecord).OldMove = QueryResults("RBYCompatible")
        Moves(CurrentRecord).SelfMove = QueryResults("AffectsSelf")
        If Len(QueryResults("RBYTM")) > 0 Then
            Moves(CurrentRecord).OldTM = QueryResults("RBYTM")
        End If
        If Len(QueryResults("GSTM")) > 0 Then
            Moves(CurrentRecord).NewTM = QueryResults("GSTM")
        End If
        Moves(CurrentRecord).SubstituteBlocks = QueryResults("BlockSubstitute")
        QueryResults.MoveNext
    Wend
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM BattleChart WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    While Not QueryResults.EOF
        CurrentRecord = QueryResults("ID")
        For X = 1 To 17
            BattleMatrix(CurrentRecord, X) = QueryResults(X)
        Next
        QueryResults.MoveNext
    Wend
    DBChecksum = GetDBChecksum
    'MainContainer.MousePointer = vbNormal
End Sub

Sub ScanForDuplicates(ByVal PKMN As Integer)
    'Clean up duplicate moves in the database
    Dim X As Integer
    Dim Y As Integer
    
    'Move Tutor vs. All (except Level)
    For X = 1 To 3
        If BasePKMN(PKMN).MoveTutor(X) > 0 Then
            For Y = 1 To 251
                If BasePKMN(PKMN).MachineMoves(Y) = BasePKMN(PKMN).MoveTutor(X) Then BasePKMN(PKMN).MachineMoves(Y) = BasePKMN(PKMN).MachineMoves(Y) * -1
                If BasePKMN(PKMN).BreedingMoves(Y) = BasePKMN(PKMN).MoveTutor(X) Then BasePKMN(PKMN).BreedingMoves(Y) = BasePKMN(PKMN).BreedingMoves(Y) * -1
                If BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).MoveTutor(X) Then BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).RBYMoves(Y) * -1
                If BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).MoveTutor(X) Then BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).SpecialMoves(Y) * -1
            Next
        End If
    Next
    
    'Base vs. All
    For X = 1 To 251
        If BasePKMN(PKMN).BaseMoves(X) > 0 Then
            For Y = 1 To 251
                If BasePKMN(PKMN).MachineMoves(Y) = BasePKMN(PKMN).BaseMoves(X) Then BasePKMN(PKMN).MachineMoves(Y) = BasePKMN(PKMN).MachineMoves(Y) * -1
                If BasePKMN(PKMN).BreedingMoves(Y) = BasePKMN(PKMN).BaseMoves(X) Then BasePKMN(PKMN).BreedingMoves(Y) = BasePKMN(PKMN).BreedingMoves(Y) * -1
                If BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).BaseMoves(X) Then BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).RBYMoves(Y) * -1
                If BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).BaseMoves(X) Then BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).SpecialMoves(Y) * -1
                If Y >= 1 And Y <= 3 Then
                    If BasePKMN(PKMN).MoveTutor(Y) = BasePKMN(PKMN).BaseMoves(X) Then BasePKMN(PKMN).MoveTutor(Y) = BasePKMN(PKMN).MoveTutor(Y) * -1
                End If
            Next
        End If
    Next
    
    'Machine vs. All
    For X = 1 To 251
        If BasePKMN(PKMN).MachineMoves(X) > 0 Then
            For Y = 1 To 251
                If BasePKMN(PKMN).BreedingMoves(Y) = BasePKMN(PKMN).MachineMoves(X) Then BasePKMN(PKMN).BreedingMoves(Y) = BasePKMN(PKMN).BreedingMoves(Y) * -1
                If BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).MachineMoves(X) Then BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).RBYMoves(Y) * -1
                If BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).MachineMoves(X) Then BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).SpecialMoves(Y) * -1
            Next
        End If
    Next
    
    'Breeding/RBGY
    For X = 1 To 251
        If BasePKMN(PKMN).BreedingMoves(X) > 0 Then
            For Y = 1 To 251
                If BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).BreedingMoves(X) Then BasePKMN(PKMN).RBYMoves(Y) = BasePKMN(PKMN).RBYMoves(Y) * -1
                If BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).BreedingMoves(X) Then BasePKMN(PKMN).SpecialMoves(Y) = BasePKMN(PKMN).SpecialMoves(Y) * -1
            Next
        End If
    Next
End Sub

Function LegalMove(ByVal Pokemon As Integer, ByVal MoveArray As Variant, ByVal AtkDV As Integer, ByVal DefDV As Integer, ByVal SpdDV As Integer, ByVal SpcDV As Integer) As String
    Dim X As Integer
    Dim Y As Integer
    Dim BreedingMoves As Integer
    Dim GSBreedingMoves As Integer
    Dim RBYMoves As Integer
    Dim InvalidMove As Boolean
    Dim SpecialMoves As Integer
    Dim GSSpecialMoves As Integer
    Dim SurfingPika As Integer
    Dim OddEggPoke As Boolean
    Dim Sunflora1 As Boolean
    Dim Sunflora2 As Boolean
    Dim Sunflora3 As Boolean
    Dim Sunflora4 As Boolean
    Dim Mv(4) As Integer
    For X = 1 To 4
        Mv(X) = CInt(MoveArray(X))
    Next X
    Sunflora1 = False
    Sunflora2 = False
    Sunflora3 = False
    Sunflora4 = False
    InvalidMove = False
    BreedingMoves = 0
    GSBreedingMoves = 0
    RBYMoves = 0
    LegalMove = ""
    
    For X = 1 To 4
        If Mv(X) > 0 Then
            For Y = 1 To 251
                If BasePKMN(Pokemon).BreedingMoves(Y) = Mv(X) Then
                    BreedingMoves = BreedingMoves + 1
                End If
                If BasePKMN(Pokemon).BreedingMoves(Y) = Mv(X) And Moves(Mv(X)).OldMove = False Then
                    GSBreedingMoves = GSBreedingMoves + 1
                End If
            Next
            If Pokemon = 207 And Mv(X) = 248 Then
                BreedingMoves = BreedingMoves - 1
            End If
        End If
    Next
    
    If Pokemon = 192 Then
        For X = 1 To 4
            If Mv(X) = 117 Then Sunflora1 = True
            If Mv(X) = 159 Then Sunflora2 = True
            If Mv(X) = 223 Then Sunflora3 = True
            If Mv(X) = 141 Then Sunflora4 = True
        Next
        If Sunflora1 And Sunflora2 Then
            LegalMove = "Sunflora cannot have Mega Drain and Razor Leaf."
            Exit Function
        End If
        If Sunflora3 And Sunflora4 Then
            LegalMove = "Sunflora cannot have Synthesis and Petal Dance."
            Exit Function
        End If
    End If
    
    'This bit is commented out until I have more info.
    'If 4 - BasePKMN(Pokemon).StartsWith < BreedingMoves Then
    '    MsgBox BasePKMN(Pokemon).Name & " can have a maximum of " & 4 - BasePKMN(Pokemon).StartsWith & " breeding moves.", vbCritical, "Illegal Move"
    '    LegalMove = False
    '    Exit Function
    'End If
    
    For X = 1 To 4
        If Mv(X) > 0 Then
            For Y = 1 To 251
                If BasePKMN(Pokemon).RBYMoves(Y) = Mv(X) Then
                    RBYMoves = RBYMoves + 1
                End If
            Next
        End If
    Next
    
    If Pokemon = 172 Or 25 Or 26 Or 173 Or 35 Or 36 Or 174 Or 39 Or 40 Or 236 Or 106 Or 107 Or 237 Or 238 Or 124 Or 239 Or 125 Or 240 Or 126 Then
        For X = 1 To 4
            If Mv(X) = 45 Then OddEggPoke = True
        Next
    End If
    
    If OddEggPoke And GSBreedingMoves > 0 Then
        LegalMove = "Sorry, no G/S/C breeding moves on a Dizzy Punchin' " & BasePKMN(Pokemon).Name
        Exit Function
    End If
    
    If OddEggPoke And Not ((AtkDV = 2 And DefDV = 10 And SpcDV = 10 And SpdDV = 10) Or (AtkDV = 0 And DefDV = 0 And SpcDV = 0 And SpdDV = 0)) Then
        LegalMove = "DVs must be either 2/10/10/10 or 0/0/0/0 in order for Dizzy Punch to be on " & BasePKMN(Pokemon).Name
        Exit Function
    End If

    If GSBreedingMoves > 0 And RBYMoves > 0 Then
        LegalMove = "Can't combine R/B/Y moves and non-R/B/Y breeding moves on " & BasePKMN(Pokemon).Name
        Exit Function
    End If
    
    For X = 1 To 4
        If Mv(X) > 0 Then
            For Y = 1 To 251
                If BasePKMN(Pokemon).SpecialMoves(Y) = Mv(X) Then
                    SpecialMoves = SpecialMoves + 1
                    If (Pokemon = 25 Or Pokemon = 26) And Mv(X) = 217 Then
                        SpecialMoves = SpecialMoves - 1
                    End If
                End If
                If BasePKMN(Pokemon).SpecialMoves(Y) = Mv(X) And Moves(Mv(X)).OldMove = False Then
                    GSSpecialMoves = GSSpecialMoves + 1
                End If
            Next
        End If
    Next

    If SpecialMoves > 1 Then
        If Pokemon = 25 Or Pokemon = 26 Then
            LegalMove = BasePKMN(Pokemon).Name & " can only have one Special move, not counting Surf."
        Else
            LegalMove = BasePKMN(Pokemon).Name & " has " & SpecialMoves & " Special moves, you can only have one Special move per Pokémon."
        End If
        Exit Function
    End If
    
    If SpecialMoves > 0 And BreedingMoves > 0 Then
        If Pokemon = 207 Then
            LegalMove = "Can't combine Breeding and Special moves on " & BasePKMN(Pokemon).Name & ", except for Earthquake and Wing Attack."
            Exit Function
        Else
            LegalMove = "Can't combine Breeding and Special moves on " & BasePKMN(Pokemon).Name
            Exit Function
        End If
    End If
    
    If Pokemon = 25 Or Pokemon = 26 Then
        For X = 1 To 4
            If Mv(X) = 217 Then
                SurfingPika = True
            End If
        Next
        If SurfingPika And GSBreedingMoves > 0 Then
            LegalMove = "Can't have G/S breeding moves on a Surfing " & BasePKMN(Pokemon).Name
            Exit Function
        End If
        If SurfingPika And GSSpecialMoves > 0 Then
            LegalMove = "Can't have G/S Special moves on a Surfing " & BasePKMN(Pokemon).Name
            Exit Function
        End If
    End If
    
    For X = 1 To 4
        If Mv(X) > 0 Then
            InvalidMove = True
            For Y = 1 To 251
                If Abs(BasePKMN(Pokemon).BaseMoves(Y)) = Mv(X) Then InvalidMove = False
                If Abs(BasePKMN(Pokemon).MachineMoves(Y)) = Mv(X) Then InvalidMove = False
                If Abs(BasePKMN(Pokemon).BreedingMoves(Y)) = Mv(X) Then InvalidMove = False
                If Abs(BasePKMN(Pokemon).RBYMoves(Y)) = Mv(X) Then InvalidMove = False
                If Abs(BasePKMN(Pokemon).SpecialMoves(Y)) = Mv(X) Then InvalidMove = False
                If Y <= 3 Then
                    If Abs(BasePKMN(Pokemon).MoveTutor(Y)) = Mv(X) Then InvalidMove = False
                End If
            Next
            If InvalidMove Then
                LegalMove = BasePKMN(Pokemon).Name & " can't learn " & Moves(Mv(X)).Name & " - there may have been a recent change to the database."
                Exit Function
            End If
        End If
    Next
End Function

